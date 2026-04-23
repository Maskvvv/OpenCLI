#!/usr/bin/env node
'use strict';

const fs = require('node:fs');
const path = require('node:path');
const readline = require('node:readline');
const { spawn, spawnSync } = require('node:child_process');

const rootDir = path.resolve(__dirname, '..');
const bundledNode = process.platform === 'win32'
  ? path.join(rootDir, 'node.exe')
  : path.join(rootDir, 'node');
const nodeBin = fs.existsSync(bundledNode) ? bundledNode : process.execPath;
const opencliEntry = path.join(rootDir, 'dist', 'src', 'main.js');
const exportsDir = path.join(rootDir, 'exports');

function createRl() {
  return readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
}

function ask(rl, question) {
  return new Promise((resolve) => rl.question(question, resolve));
}

function stripAnsi(text) {
  return String(text || '').replace(/\x1B\[[0-?]*[ -/]*[@-~]/g, '');
}

async function askPositiveInt(rl, label, fallback) {
  while (true) {
    const suffix = fallback === undefined ? '' : `，默认 ${fallback}`;
    const raw = (await ask(rl, `${label}${suffix}: `)).trim();
    const value = raw || String(fallback ?? '');
    const parsed = Number.parseInt(value, 10);
    if (Number.isInteger(parsed) && parsed >= 1) {
      return parsed;
    }
    console.log('请输入大于等于 1 的整数。');
  }
}

async function askOutputPath(rl, startPage, endPage) {
  const defaultOutput = path.join(exportsDir, `pgy-${startPage}-${endPage}.csv`);
  const raw = (await ask(rl, `输出 CSV 路径，直接回车使用默认值 (${defaultOutput}): `)).trim();
  if (!raw) return defaultOutput;
  return path.resolve(rootDir, raw);
}

function runOpenCli(args, options = {}) {
  if (!fs.existsSync(opencliEntry)) {
    throw new Error(`找不到 OpenCLI 入口文件: ${opencliEntry}`);
  }
  return spawnSync(nodeBin, [opencliEntry, ...args], {
    cwd: rootDir,
    env: {
      ...process.env,
      OPENCLI_BROWSER_CONNECT_TIMEOUT: process.env.OPENCLI_BROWSER_CONNECT_TIMEOUT || '20',
    },
    encoding: 'utf-8',
    windowsHide: false,
    ...options,
  });
}

function printExtensionGuide() {
  console.log('');
  console.log('没有检测到 OpenCLI Chrome 扩展连接。请按下面步骤处理：');
  console.log(`1. 打开 Chrome，进入 chrome://extensions/`);
  console.log('2. 打开右上角“开发者模式”。');
  console.log(`3. 点击“加载已解压的扩展程序”，选择这个文件夹: ${path.join(rootDir, 'extension')}`);
  console.log('4. 确认 OpenCLI 扩展已启用。');
  console.log('5. 在 Chrome 中登录 https://pgy.xiaohongshu.com/ 后回到这里。');
  console.log('');
}

async function ensureExtensionConnected(rl) {
  while (true) {
    console.log('正在检查 OpenCLI daemon 和 Chrome 扩展连接状态...');
    const result = runOpenCli(['doctor', '--no-live']);
    const output = `${result.stdout || ''}${result.stderr || ''}`;
    const cleanOutput = stripAnsi(output);
    if (/\[OK\]\s*Extension:\s*connected/i.test(cleanOutput)) {
      console.log('Chrome 扩展已连接。');
      return;
    }

    if (cleanOutput.trim()) {
      console.log('');
      console.log(cleanOutput.trim());
    }
    printExtensionGuide();
    await ask(rl, '完成后按 Enter 重新检查，或按 Ctrl+C 退出。');
  }
}

async function main() {
  console.log('OpenCLI 蒲公英 pgycsv 导出工具');
  console.log('请先确保 Chrome 已安装、OpenCLI 扩展已启用，并已登录蒲公英。');
  console.log('');

  fs.mkdirSync(exportsDir, { recursive: true });
  const rl = createRl();
  try {
    await ensureExtensionConnected(rl);
    console.log('');
    const startPage = await askPositiveInt(rl, '从第几页开始', 1);
    const endPage = await askPositiveInt(rl, '到第几页结束', startPage);
    if (endPage < startPage) {
      console.log('结束页不能小于起始页，请重新运行。');
      process.exitCode = 1;
      return;
    }
    const outputPath = await askOutputPath(rl, startPage, endPage);
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });

    console.log('');
    console.log(`开始导出第 ${startPage} 页到第 ${endPage} 页...`);
    console.log(`CSV 将保存到: ${outputPath}`);
    console.log('');

    const child = spawn(nodeBin, [
      opencliEntry,
      'xiaohongshu',
      'pgycsv',
      '--start-page',
      String(startPage),
      '--end-page',
      String(endPage),
      '--output',
      outputPath,
    ], {
      cwd: rootDir,
      env: {
        ...process.env,
        OPENCLI_WINDOW_FOCUSED: process.env.OPENCLI_WINDOW_FOCUSED || '1',
      },
      stdio: 'inherit',
      windowsHide: false,
    });

    const exitCode = await new Promise((resolve) => {
      child.on('close', resolve);
      child.on('error', (err) => {
        console.error(`启动导出失败: ${err.message}`);
        resolve(1);
      });
    });

    process.exitCode = Number(exitCode || 0);
    console.log('');
    if (process.exitCode === 0) {
      console.log(`导出完成: ${outputPath}`);
    } else {
      console.log('导出未成功。请查看上方错误信息，确认 Chrome 扩展已连接且蒲公英已登录。');
    }
  } finally {
    rl.close();
  }
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exitCode = 1;
});
