#!/usr/bin/env node
'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const rootDir = path.resolve(__dirname, '..');
const releaseRoot = path.join(rootDir, 'release');
const packageDir = path.join(releaseRoot, 'pgycsv-windows');

function run(command, args, options = {}) {
  console.log(`> ${command} ${args.join(' ')}`);
  const result = spawnSync(command, args, {
    cwd: rootDir,
    stdio: 'inherit',
    ...options,
  });
  if (result.error) {
    console.error(result.error.message);
  }
  if (result.status !== 0) {
    process.exit(result.status || 1);
  }
}

function runNpmScript(name) {
  if (process.platform === 'win32') {
    run('cmd.exe', ['/d', '/s', '/c', `npm.cmd run ${name}`]);
    return;
  }
  run('npm', ['run', name]);
}

function copyDir(src, dest, options = {}) {
  const { filter } = options;
  fs.cpSync(src, dest, {
    recursive: true,
    force: true,
    filter: (source) => {
      const rel = path.relative(src, source).replace(/\\/g, '/');
      return filter ? filter(rel, source) : true;
    },
  });
}

function copyFile(src, dest) {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

function writeText(filePath, content) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, 'utf-8');
}

function ensureExists(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${label} not found: ${filePath}`);
  }
}

function copyExtension() {
  const src = path.join(rootDir, 'extension');
  const dest = path.join(packageDir, 'extension');
  ensureExists(path.join(src, 'manifest.json'), 'extension manifest');
  ensureExists(path.join(src, 'dist', 'background.js'), 'extension background');
  copyDir(src, dest, {
    filter: (rel) => {
      if (!rel) return true;
      if (rel === 'node_modules' || rel.startsWith('node_modules/')) return false;
      if (rel === 'src' || rel.startsWith('src/')) return false;
      if (rel === 'scripts' || rel.startsWith('scripts/')) return false;
      if (rel === 'store-assets' || rel.startsWith('store-assets/')) return false;
      if (rel.endsWith('.ts') || rel.endsWith('.lock')) return false;
      if (rel === 'package.json' || rel === 'package-lock.json' || rel === 'bun.lock' || rel === 'tsconfig.json' || rel === 'vite.config.ts') return false;
      return true;
    },
  });
}

function writeLauncherFiles() {
  copyFile(
    path.join(rootDir, 'tools', 'pgycsv-windows-launcher.cjs'),
    path.join(packageDir, 'tools', 'pgycsv-windows-launcher.cjs'),
  );

  writeText(path.join(packageDir, '启动 pgycsv.bat'), [
    '@echo off',
    'setlocal',
    'cd /d "%~dp0"',
    'if not exist "node.exe" (',
    '  echo Cannot find bundled node.exe.',
    '  pause',
    '  exit /b 1',
    ')',
    '"%~dp0node.exe" "%~dp0tools\\pgycsv-windows-launcher.cjs"',
    'echo.',
    'pause',
    '',
  ].join('\r\n'));

  writeText(path.join(packageDir, '安装说明.txt'), [
    'OpenCLI 蒲公英 pgycsv Windows 便携包',
    '',
    '使用前准备：',
    '1. 目标电脑需要安装 Chrome。',
    '2. 打开 Chrome，进入 chrome://extensions/。',
    '3. 打开右上角“开发者模式”。',
    '4. 点击“加载已解压的扩展程序”。',
    '5. 选择本便携包里的 extension 文件夹。',
    '6. 确认 OpenCLI 扩展已启用。',
    '7. 在 Chrome 中打开 https://pgy.xiaohongshu.com/ 并登录。',
    '',
    '开始导出：',
    '1. 双击“启动 pgycsv.bat”。',
    '2. 按提示输入起始页、结束页和 CSV 输出路径。',
    '3. 如果输出路径直接回车，CSV 会保存到 exports 文件夹。',
    '',
    '提示：',
    '- 运行时终端会打印每个博主的数据行，便于检查。',
    '- 如果提示扩展未连接，请确认 Chrome 已打开、扩展已启用，然后回到终端按 Enter 重试。',
    '- 本包已内置 Node 和 OpenCLI，不需要在目标电脑安装 Node、npm 或 OpenCLI。',
    '',
  ].join('\r\n'));
}

function main() {
  runNpmScript('build');

  fs.rmSync(packageDir, { recursive: true, force: true });
  fs.mkdirSync(packageDir, { recursive: true });

  copyFile(process.execPath, path.join(packageDir, process.platform === 'win32' ? 'node.exe' : 'node'));
  copyDir(path.join(rootDir, 'dist'), path.join(packageDir, 'dist'));
  copyDir(path.join(rootDir, 'clis'), path.join(packageDir, 'clis'));
  copyDir(path.join(rootDir, 'node_modules'), path.join(packageDir, 'node_modules'), {
    filter: (rel) => !(rel === '.cache' || rel.startsWith('.cache/')),
  });
  copyFile(path.join(rootDir, 'package.json'), path.join(packageDir, 'package.json'));
  copyFile(path.join(rootDir, 'cli-manifest.json'), path.join(packageDir, 'cli-manifest.json'));
  copyExtension();
  writeLauncherFiles();
  fs.mkdirSync(path.join(packageDir, 'exports'), { recursive: true });

  console.log('');
  console.log(`pgycsv Windows portable package created: ${packageDir}`);
  console.log('You can zip this folder and copy it to another Windows computer.');
}

main();
