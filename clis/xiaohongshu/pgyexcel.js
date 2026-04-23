/**
 * Xiaohongshu Pugongying invite-list exporter.
 *
 * Visits the Pugongying brand invite list, filters interested creators,
 * opens each creator detail, and writes Excel-friendly CSV rows.
 *
 * Requires: logged into pgy.xiaohongshu.com in Chrome.
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { cli, Strategy } from '@jackwener/opencli/registry';
import { ArgumentError, AuthRequiredError, EmptyResultError } from '@jackwener/opencli/errors';

export const PGY_INVITE_LIST_URL = 'https://pgy.xiaohongshu.com/solar/pre-trade/brand/invite-list/note';
export const CSV_COLUMNS = ['联系方式', '达人昵称', '达人ID', '达人类型1', '达人类型2', '蒲公英连接', '主页链接'];

const DETAIL_SETTLE_SECONDS = 2.5;
const LIST_READY_TIMEOUT_MS = 15_000;
const TABLE_READY_TIMEOUT_MS = 25_000;

function cleanText(value) {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function clampDelayBounds(minDelay, maxDelay) {
    const min = Number(minDelay ?? 1.2);
    const max = Number(maxDelay ?? 3.8);
    if (!Number.isFinite(min) || !Number.isFinite(max) || min < 0 || max < 0) {
        throw new ArgumentError('--min-delay and --max-delay must be non-negative numbers');
    }
    if (max < min) {
        throw new ArgumentError('--max-delay must be greater than or equal to --min-delay');
    }
    return { min, max };
}

function randomDelaySeconds(min, max) {
    if (max <= min)
        return min;
    return min + Math.random() * (max - min);
}

async function humanPause(page, delayBounds, multiplier = 1) {
    await page.wait({ time: randomDelaySeconds(delayBounds.min, delayBounds.max) * multiplier });
}

export function normalizePageRange(kwargs) {
    const startPage = Math.trunc(Number(kwargs['start-page'] ?? 1));
    const endRaw = kwargs['end-page'] ?? startPage;
    const endPage = Math.trunc(Number(endRaw));
    if (!Number.isFinite(startPage) || startPage < 1) {
        throw new ArgumentError('--start-page must be an integer greater than or equal to 1');
    }
    if (!Number.isFinite(endPage) || endPage < startPage) {
        throw new ArgumentError('--end-page must be an integer greater than or equal to --start-page');
    }
    return { startPage, endPage };
}

export function escapeCsvCell(value) {
    const text = String(value ?? '');
    const escaped = text.replace(/"/g, '""');
    return /[",\r\n]/.test(escaped) ? `"${escaped}"` : escaped;
}

export function rowsToCsv(rows) {
    const lines = [
        CSV_COLUMNS.map(escapeCsvCell).join(','),
        ...rows.map((row) => CSV_COLUMNS.map((column) => escapeCsvCell(row[column])).join(',')),
    ];
    return `\uFEFF${lines.join('\n')}\n`;
}

export function isInterestedInviteText(value) {
    const text = cleanText(value);
    return text.includes('感兴趣') && !text.includes('不感兴趣');
}

export function writeCsv(outputPath, rows) {
    const resolved = path.resolve(String(outputPath || './xiaohongshu-pgy-creators.csv'));
    fs.mkdirSync(path.dirname(resolved), { recursive: true });
    fs.writeFileSync(resolved, rowsToCsv(rows), 'utf-8');
    return resolved;
}

export function parseCreatorProfileFromDomData(data) {
    const text = cleanText(data?.noteHomeText || data?.bodyText || '');
    const labels = Array.isArray(data?.labels)
        ? [...new Set(data.labels
            .flatMap((label) => cleanText(label).split(/[、\s]+/))
            .map(cleanText)
            .filter((label) => label && label.length <= 12 && !/^\d/.test(label)))]
        : [];
    const links = Array.isArray(data?.links) ? data.links : [];

    const nickname = cleanText(data?.fallbackNickname)
        || cleanText(data?.nickname)
        || cleanText(text.match(/达人昵称[:：]?\s*([^\s|]+)/)?.[1])
        || cleanText(text.match(/昵称[:：]?\s*([^\s|]+)/)?.[1]);

    const creatorId = cleanText(data?.creatorId)
        || cleanText(text.match(/小红书号[:：]?\s*([A-Za-z0-9._-]+)/)?.[1])
        || cleanText(text.match(/达人ID[:：]?\s*([A-Za-z0-9._-]+)/)?.[1]);

    const homepageUrl = cleanText(data?.homepageUrl)
        || cleanText(data?.fallbackHomepageUrl)
        || cleanText(links.find((href) => /xiaohongshu\.com\/user\/profile\//.test(String(href))) || '');

    return {
        '联系方式': cleanText(data?.contact),
        '达人昵称': nickname,
        '达人ID': creatorId,
        '达人类型1': labels[0] ?? '',
        '达人类型2': labels[1] ?? '',
        '蒲公英连接': cleanText(data?.pgyUrl),
        '主页链接': homepageUrl,
    };
}

export function extractInviteRowsFromDomData(items) {
    if (!Array.isArray(items))
        return [];
    return items
        .map((item, index) => ({
        index,
        nickname: cleanText(item.nickname),
        detailText: cleanText(item.detailText),
        hasDetailButton: Boolean(item.hasDetailButton),
        hasProfileTarget: Boolean(item.hasProfileTarget),
        }))
        .filter((item) => isInterestedInviteText(item.detailText))
        .filter((item) => item.hasDetailButton || item.hasProfileTarget || item.nickname);
}

async function ensurePgyReady(page) {
    const status = await page.evaluate(`
    new Promise((resolve) => {
      const detect = () => {
        const href = location.href;
        const bodyText = document.body?.innerText || '';
        if (/login|passport|signin/i.test(href) || /登录|扫码|验证码/.test(bodyText)) return 'login';
        if (/我的邀约|感兴趣|查看详情/.test(bodyText)) return 'ready';
        return '';
      };
      const current = detect();
      if (current) return resolve(current);
      const observer = new MutationObserver(() => {
        const next = detect();
        if (next) {
          observer.disconnect();
          resolve(next);
        }
      });
      observer.observe(document.body, { childList: true, subtree: true });
      setTimeout(() => {
        observer.disconnect();
        resolve(detect() || 'timeout');
      }, ${LIST_READY_TIMEOUT_MS});
    })
  `);
    if (status === 'login') {
        throw new AuthRequiredError('pgy.xiaohongshu.com', 'Pugongying invite list requires login');
    }
    if (status !== 'ready') {
        throw new EmptyResultError('xiaohongshu/pgyexcel', 'Pugongying invite list did not load expected controls');
    }
}

async function clickText(page, labels, options = {}) {
    const labelsJson = JSON.stringify(labels);
    const exactJson = JSON.stringify(Boolean(options.exact));
    const result = await page.evaluate(`
    (() => {
      const labels = ${labelsJson};
      const exact = ${exactJson};
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const normalize = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const selector = 'button, [role="button"], [role="tab"], a, label, span, div, li';
      for (const node of Array.from(document.querySelectorAll(selector))) {
        if (!isVisible(node)) continue;
        const text = normalize(node.innerText || node.textContent || '');
        if (!text) continue;
        const matched = labels.some((label) => exact ? text === label : text === label || text.includes(label));
        if (!matched) continue;
        const clickable = node.closest('button, [role="button"], [role="tab"], a, label, li') || node;
        clickable.scrollIntoView({ block: 'center', inline: 'center' });
        const rect = clickable.getBoundingClientRect();
        return {
          ok: true,
          text,
          x: Math.round(rect.left + rect.width / 2),
          y: Math.round(rect.top + rect.height / 2),
        };
      }
      return { ok: false };
    })()
  `);
    if (!result?.ok)
        return false;
    if (page.nativeClick && Number.isFinite(result.x) && Number.isFinite(result.y)) {
        await page.nativeClick(result.x, result.y);
    }
    else {
        await page.evaluate(`() => document.elementFromPoint(${Number(result.x)}, ${Number(result.y)})?.click()`);
    }
    return true;
}

async function clickElementByEval(page, jsLocator) {
    const target = await page.evaluate(`
    (() => {
      const el = (${jsLocator})();
      if (!el) return { ok: false };
      el.scrollIntoView({ block: 'center', inline: 'center' });
      const rect = el.getBoundingClientRect();
      return { ok: true, x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
    })()
  `);
    if (!target?.ok)
        return false;
    if (page.nativeClick) {
        await page.nativeClick(target.x, target.y);
    }
    else {
        await page.evaluate(`() => document.elementFromPoint(${Number(target.x)}, ${Number(target.y)})?.click()`);
    }
    return true;
}

async function applyInterestedInviteFilter(page, delayBounds) {
    await clickText(page, ['我的邀约']);
    await humanPause(page, delayBounds);
    const clickedFilter = await clickElementByEval(page, `
    () => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      return Array.from(document.querySelectorAll('button, [role="tab"], [role="button"]'))
        .find((el) => clean(el.innerText || el.textContent || '').startsWith('感兴趣')) || null;
    }
  `);
    if (!clickedFilter) {
        await clickText(page, ['邀约状态', '状态', '筛选', '全部']);
        await humanPause(page, delayBounds, 0.7);
        await clickText(page, ['感兴趣']);
    }
    await humanPause(page, delayBounds, 0.5);
    await clickText(page, ['查询'], { exact: true });
    await humanPause(page, delayBounds);
    await settleInterestedInviteFilter(page, delayBounds);
}

async function waitForInviteTable(page) {
    await page.evaluate(`
    new Promise((resolve) => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const hasRows = () => {
        const bodyText = clean(document.body?.innerText || '');
        if (/暂无数据|暂无结果|无数据/.test(bodyText)) return true;
        const candidates = Array.from(document.querySelectorAll('${rowSelector()}'));
        return candidates.some((el) => {
          const text = clean(el.innerText || el.textContent || '');
          return visible(el) && text.includes('查看详情');
        });
      };
      if (hasRows()) return resolve(true);
      const observer = new MutationObserver(() => {
        if (hasRows()) {
          observer.disconnect();
          resolve(true);
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      setTimeout(() => {
        observer.disconnect();
        resolve(true);
      }, ${TABLE_READY_TIMEOUT_MS});
    })
  `);
}

async function waitForInterestedInviteRows(page, timeoutMs = TABLE_READY_TIMEOUT_MS) {
    return page.evaluate(`
    new Promise((resolve) => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const detect = () => {
        const rows = Array.from(document.querySelectorAll('${rowSelector()}')).filter((el) => {
          const text = clean(el.innerText || el.textContent || '');
          return visible(el) && text.includes('查看详情') && text.includes('感兴趣') && !text.includes('不感兴趣');
        });
        if (rows.length) return 'rows';
        const bodyText = clean(document.body?.innerText || '');
        if (/暂无数据|暂无结果|无数据/.test(bodyText)) return 'empty';
        return '';
      };
      const current = detect();
      if (current) return resolve(current);
      const observer = new MutationObserver(() => {
        const next = detect();
        if (next) {
          observer.disconnect();
          resolve(next);
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      setTimeout(() => {
        observer.disconnect();
        resolve(detect() || 'timeout');
      }, ${Number(timeoutMs)});
    })
  `);
}

async function settleInterestedInviteFilter(page, delayBounds) {
    for (let attempt = 0; attempt < 3; attempt++) {
        const state = await waitForInterestedInviteRows(page, 12_000);
        if (state === 'rows')
            return true;
        await clickText(page, ['查询'], { exact: true });
        await humanPause(page, delayBounds, 1.2 + attempt * 0.4);
    }
    return false;
}

function rowSelector() {
    return 'tbody tr';
}

async function readInviteTableDebug(page) {
    return page.evaluate(`
    (() => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const rows = Array.from(document.querySelectorAll('${rowSelector()}'))
        .filter(visible)
        .map((el) => clean(el.innerText || el.textContent || ''))
        .filter(Boolean)
        .filter((text) => /查看详情|感兴趣|不感兴趣|博主信息|博主意向/.test(text))
        .slice(0, 8);
      return {
        href: location.href,
        hasQuery: clean(document.body?.innerText || '').includes('查询'),
        hasInterested: clean(document.body?.innerText || '').includes('感兴趣'),
        hasDetail: clean(document.body?.innerText || '').includes('查看详情'),
        samples: rows,
      };
    })()
  `);
}

async function gotoListPageNumber(page, pageNum, delayBounds) {
    if (pageNum <= 1)
        return;
    const clickedTargetPage = await clickVisibleListPageNumber(page, pageNum);
    if (clickedTargetPage) {
        await humanPause(page, delayBounds, 1.2);
        const currentPage = await readCurrentListPageNumber(page);
        if (!currentPage || currentPage === pageNum)
            return;
    }
    for (let attempts = 0; attempts < pageNum + 4; attempts++) {
        const currentPage = await readCurrentListPageNumber(page);
        if (currentPage === pageNum)
            return;
        if (currentPage && currentPage > pageNum)
            break;
        const clicked = await clickPaginationNext(page);
        if (!clicked)
            break;
        await humanPause(page, delayBounds, 1.2);
    }
    const inputOk = await jumpListPageByInput(page, pageNum);
    if (inputOk) {
        await humanPause(page, delayBounds, 1.2);
        const currentPage = await readCurrentListPageNumber(page);
        if (!currentPage || currentPage === pageNum)
            return;
    }
    throw new EmptyResultError('xiaohongshu/pgyexcel', `Could not navigate to page ${pageNum}`);
}

async function clickVisibleListPageNumber(page, pageNum) {
    const target = await page.evaluate(`
    (() => {
      const pageNum = String(${Number(pageNum)});
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const pageTextPattern = /共\\s*\\d+\\s*条[，,]?\\s*\\d+\\s*页|跳至\\s*\\d*\\s*页|条\\/页/;
      const candidates = Array.from(document.querySelectorAll('span, div, button, a, [role="button"]'))
        .filter(visible)
        .filter((el) => clean(el.innerText || el.textContent || '') === pageNum)
        .map((el) => {
          let ancestor = el;
          let context = '';
          for (let depth = 0; ancestor && depth < 8; depth++) {
            context = clean(ancestor.innerText || ancestor.textContent || '');
            if (pageTextPattern.test(context)) break;
            ancestor = ancestor.parentElement;
          }
          const rect = el.getBoundingClientRect();
          return {
            el,
            hasPaginationContext: pageTextPattern.test(context),
            top: rect.top,
            left: rect.left,
          };
        })
        .filter((item) => item.hasPaginationContext || item.top > window.innerHeight * 0.45)
        .sort((a, b) => (Number(b.hasPaginationContext) - Number(a.hasPaginationContext)) || b.top - a.top || a.left - b.left);
      const item = candidates[0];
      if (!item) return { ok: false };
      const clickable = item.el.closest('button, a, [role="button"]') || item.el.parentElement || item.el;
      clickable.scrollIntoView({ block: 'center', inline: 'center' });
      const rect = clickable.getBoundingClientRect();
      return { ok: true, x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
    })()
  `);
    if (!target?.ok)
        return false;
    if (page.nativeClick) {
        await page.nativeClick(target.x, target.y);
    }
    else {
        await page.evaluate(`() => document.elementFromPoint(${Number(target.x)}, ${Number(target.y)})?.click()`);
    }
    return true;
}

async function readCurrentListPageNumber(page) {
    return page.evaluate(`
    (() => {
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const selectors = [
        '.ant-pagination-item-active',
        '.d-pagination-item-active',
        '.semi-page-item-active',
        '[class*="pagination"] [class*="active"]',
        '[class*="pager"] [class*="active"]',
      ];
      for (const selector of selectors) {
        for (const el of Array.from(document.querySelectorAll(selector))) {
          if (!visible(el)) continue;
          const text = clean(el.innerText || el.textContent || el.getAttribute('title') || '');
          const match = text.match(/^\\d+$/) || clean(el.getAttribute('aria-current') || '').match(/^page\\s+(\\d+)$/i);
          if (match) return Number(match[1] || match[0]);
        }
      }
      return 0;
    })()
  `);
}

async function clickPaginationNext(page) {
    const target = await page.evaluate(`
    (() => {
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const disabled = (el) => {
        const node = el.closest('button, li, [role="button"], a, div') || el;
        const cls = String(node.className || '');
        return node.disabled || node.getAttribute('aria-disabled') === 'true' || /disabled/.test(cls);
      };
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const containers = Array.from(document.querySelectorAll('[class*="pagination"], [class*="pager"], .ant-pagination, .d-pagination, .semi-page')).filter(visible);
      const roots = containers.length ? containers : [document.body];
      const candidates = [];
      for (const root of roots) {
        candidates.push(...Array.from(root.querySelectorAll('button, li, a, [role="button"], [class*="next"]')));
      }
      for (const el of candidates) {
        if (!visible(el) || disabled(el)) continue;
        const text = clean(el.innerText || el.textContent || '');
        const attrs = clean([
          el.className,
          el.getAttribute('aria-label'),
          el.getAttribute('title'),
          el.getAttribute('data-testid'),
        ].join(' '));
        const looksNext = /next|下一页|下页|后一页|right/i.test(attrs) || /^(>|›|»)$/.test(text) || /下一页|下页/.test(text);
        if (!looksNext) continue;
        const clickable = el.closest('button, a, li, [role="button"]') || el;
        clickable.scrollIntoView({ block: 'center', inline: 'center' });
        const rect = clickable.getBoundingClientRect();
        return { ok: true, x: Math.round(rect.left + rect.width / 2), y: Math.round(rect.top + rect.height / 2) };
      }
      return { ok: false };
    })()
  `);
    if (!target?.ok)
        return false;
    if (page.nativeClick) {
        await page.nativeClick(target.x, target.y);
    }
    else {
        await page.evaluate(`() => document.elementFromPoint(${Number(target.x)}, ${Number(target.y)})?.click()`);
    }
    return true;
}

async function jumpListPageByInput(page, pageNum) {
    return page.evaluate(`
    (() => {
      const pageNum = ${Number(pageNum)};
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const roots = Array.from(document.querySelectorAll('[class*="pagination"], [class*="pager"], .ant-pagination, .d-pagination, .semi-page')).filter(visible);
      const inputs = (roots.length ? roots : [document.body]).flatMap((root) => Array.from(root.querySelectorAll('input')));
      const input = inputs.find((el) => visible(el) && (/页|page|jump|goto/i.test(el.placeholder || el.getAttribute('aria-label') || '') || el.type === 'number' || el.inputMode === 'numeric'));
      if (!input) return false;
      input.focus();
      input.value = String(pageNum);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', bubbles: true }));
      return true;
    })()
  `);
}

async function gotoListPageNumberLegacy(page, pageNum, delayBounds) {
    if (pageNum <= 1)
        return;
    for (let current = 1; current < pageNum; current++) {
        const clicked = await clickText(page, ['下一页', '下一页 >', '>'], { exact: false });
        if (!clicked) {
            const inputOk = await page.evaluate(`
        (() => {
          const pageNum = ${Number(pageNum)};
          const inputs = Array.from(document.querySelectorAll('input'));
          const input = inputs.find((el) => /页|page/i.test(el.placeholder || el.getAttribute('aria-label') || '') || el.type === 'number');
          if (!input) return false;
          input.focus();
          input.value = String(pageNum);
          input.dispatchEvent(new Event('input', { bubbles: true }));
          input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
          input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', bubbles: true }));
          return true;
        })()
      `);
            if (!inputOk)
                throw new EmptyResultError('xiaohongshu/pgyexcel', `Could not navigate to page ${pageNum}`);
            break;
        }
        await humanPause(page, delayBounds, 1.2);
    }
}

async function readInviteRows(page) {
    const rows = await page.evaluate(`
    (() => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const candidates = Array.from(document.querySelectorAll('${rowSelector()}'));
      const seen = new Set();
      const rows = [];
      for (const el of candidates) {
        if (!isVisible(el)) continue;
        const text = clean(el.innerText || el.textContent || '');
        if (!text || !text.includes('查看详情') || !text.includes('感兴趣') || text.includes('不感兴趣')) continue;
        const key = text.slice(0, 120);
        if (seen.has(key)) continue;
        seen.add(key);
        const link = el.querySelector('a[href*="pgy.xiaohongshu.com"], a[href*="/solar/"], a[href*="/profile"], a[href*="kol"]');
        const nameEl = el.querySelector('.kol-name, [class*="nick"], [class*="author"], [class*="blogger"], a');
        rows.push({
          nickname: clean(nameEl?.innerText || nameEl?.textContent || ''),
          detailText: text,
          hasDetailButton: true,
          hasProfileTarget: !!link || !!nameEl,
        });
      }
      return rows;
    })()
  `);
    return extractInviteRowsFromDomData(rows);
}

async function fetchInterestedInvites(page, pageNum) {
    const result = await page.evaluate(`
    (async () => {
      const resp = await fetch('/api/solar/invite/get_invites_overview', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          kolIntention: 1,
          inviteStatus: -1,
          kolType: 0,
          showWechat: 0,
          searchDateType: 1,
          pageNum: ${Number(pageNum)},
          pageSize: 10,
        }),
      });
      const json = await resp.json();
      if (!resp.ok || json.code !== 0 || json.success === false) {
        return { ok: false, status: resp.status, message: json.msg || 'failed to load invite list' };
      }
      return {
        ok: true,
        total: json.data?.total || 0,
        invites: (json.data?.invites || []).map((item) => ({
          inviteId: item.inviteId || '',
          kolId: item.kolId || '',
          kolName: item.kolName || '',
          kolIntention: item.kolIntention,
          kolType: item.kolType,
        })),
      };
    })()
  `);
    if (!result?.ok) {
        throw new EmptyResultError('xiaohongshu/pgyexcel', `Could not load interested invite list: ${result?.message || result?.status || 'unknown error'}`);
    }
    return (result.invites || []).filter((item) => item.kolId && Number(item.kolIntention) === 1);
}

async function readContactInfoByKolId(page, kolId) {
    if (!kolId)
        return '';
    const result = await page.evaluate(`
    (async () => {
      const kolId = ${JSON.stringify(kolId)};
      try {
        const auth = await fetch('/api/pgy/kol/contact/whitelist/view?kolId=' + encodeURIComponent(kolId)).then((resp) => resp.json());
        const cipherText = auth?.data?.cipherText;
        if (!cipherText) return '';
        const plain = await fetch('/api/pgy/kol/contact/view', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({
            kolId,
            cipherText: {
              phoneCipherText: cipherText.phoneCipherText || '',
              weChatCipherText: cipherText.weChatCipherText || '',
            },
          }),
        }).then((resp) => resp.json());
        const text = plain?.data?.plainText || plain?.plainText || plain?.data || {};
        return text.weChat || text.wechat || text.weChatPlainText || text.phone || text.tel || '';
      } catch {
        return '';
      }
    })()
  `);
    return cleanText(result);
}

async function openCreatorDetailByKolId(page, kolId, delayBounds) {
    await page.goto(`https://pgy.xiaohongshu.com/solar/pre-trade/blogger-detail/${encodeURIComponent(kolId)}`);
    await humanPause(page, delayBounds, 1.2);
    await waitForPgyDetail(page, '');
}

async function clickInviteDetailButton(page, rowIndex) {
    return clickElementByEval(page, `
    () => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const rows = Array.from(document.querySelectorAll('${rowSelector()}'))
        .filter((el) => {
          const text = clean(el.innerText || el.textContent || '');
          return isVisible(el) && text.includes('查看详情') && text.includes('感兴趣') && !text.includes('不感兴趣');
        });
      const row = rows[${Number(rowIndex)}];
      if (!row) return null;
      return Array.from(row.querySelectorAll('button, [role="button"], a, span, div'))
        .find((el) => isVisible(el) && clean(el.innerText || el.textContent || '') === '查看详情') || null;
    }
  `);
}

async function clickInviteDetailButtonForInvite(page, invite) {
    return clickElementByEval(page, `
    () => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const nickname = ${JSON.stringify(invite?.kolName || '')};
      const kolId = ${JSON.stringify(invite?.kolId || '')};
      const rows = Array.from(document.querySelectorAll('${rowSelector()}'))
        .filter((el) => {
          const text = clean(el.innerText || el.textContent || '');
          return isVisible(el) && text.includes('查看详情') && text.includes('感兴趣') && !text.includes('不感兴趣');
        });
      const row = rows.find((el) => {
        const text = clean(el.innerText || el.textContent || '');
        const html = el.innerHTML || '';
        return (nickname && text.includes(nickname)) || (kolId && html.includes(kolId));
      }) || rows[${Number(invite?.rowIndex ?? -1)}] || null;
      if (!row) return null;
      return Array.from(row.querySelectorAll('button, [role="button"], a, span, div'))
        .find((el) => isVisible(el) && clean(el.innerText || el.textContent || '') === '查看详情') || null;
    }
  `);
}

async function clickRevealContactButton(page) {
    return page.evaluate(`
    () => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const textOf = (el) => clean(el.innerText || el.textContent || '');
      const activeSurface = () => Array.from(document.querySelectorAll('.d-drawer, .ant-drawer, .ant-modal, .semi-modal, [role="dialog"], [class*="drawer"], [class*="popover"]'))
        .filter(isVisible)
        .filter((el) => {
          const rect = el.getBoundingClientRect();
          return rect.left < window.innerWidth && rect.right > 0 && /邀约详情|博主微信/.test(textOf(el));
        })
        .sort((a, b) => b.getBoundingClientRect().right - a.getBoundingClientRect().right)[0];
      const surface = activeSurface() || document.body;
      const fieldRow = Array.from(surface.querySelectorAll('.flex.gap-16, .flex, [class*="row"]'))
        .filter(isVisible)
        .filter((el) => {
          const text = textOf(el);
          return text.includes('博主微信') && text.includes('***') && el.querySelector('.item-right, [class*="right"]');
        })
        .sort((a, b) => textOf(a).length - textOf(b).length)[0];
      const icon = fieldRow
        ? Array.from(fieldRow.querySelectorAll('.cursor-animation, .d-icon, [class*="eye"], [class*="view"], button, [role="button"], a'))
          .filter(isVisible)
          .filter((el) => !/博主微信|联系方式属于|\\*\\*\\*/.test(textOf(el)))
          .sort((a, b) => b.getBoundingClientRect().left - a.getBoundingClientRect().left)[0]
        : null;
      if (icon) {
        const target = icon.closest('button, [role="button"], a') || icon;
        const rect = target.getBoundingClientRect();
        const clientX = rect.left + rect.width / 2;
        const clientY = rect.top + rect.height / 2;
        for (const type of ['pointerover', 'pointerenter', 'mouseover', 'mouseenter', 'pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click']) {
          const event = type.startsWith('pointer') && window.PointerEvent
            ? new PointerEvent(type, {
              bubbles: true,
              cancelable: true,
              view: window,
              clientX,
              clientY,
              pointerId: 1,
              pointerType: 'mouse',
              isPrimary: true,
            })
            : new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX, clientY });
          target.dispatchEvent(event);
        }
        return true;
      }
      const fallbackLabel = Array.from(surface.querySelectorAll('*'))
        .filter(isVisible)
        .find((el) => /^博主微信$/.test(textOf(el)) || /博主微信/.test(textOf(el)));
      if (!fallbackLabel) return false;
      const labelRect = fallbackLabel.getBoundingClientRect();
      const candidates = Array.from(surface.querySelectorAll('button, [role="button"], a, span, div'))
        .filter(isVisible)
        .filter((el) => {
          const text = textOf(el);
          if (!/^查看$|^查看联系方式$|^查看微信$/.test(text)) return false;
          if (textOf(el.closest('${rowSelector()}') || el).includes('查看详情')) return false;
          const rect = el.getBoundingClientRect();
          return Math.abs((rect.top + rect.height / 2) - (labelRect.top + labelRect.height / 2)) < 90
            && rect.left > labelRect.left;
        });
      const target = candidates.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)[0];
      if (!target) return false;
      target.click();
      return true;
    }
  `);
}

async function readWechatContact(page) {
    const contact = await page.evaluate(`
    (() => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
      };
      const valid = (value) => {
        const text = clean(value);
        if (!text || /^\\*+$/.test(text) || /联系方式属于|未经博主同意|联系电话/.test(text)) return '';
        return text;
      };
      const surface = Array.from(document.querySelectorAll('.d-drawer, .ant-drawer, .ant-modal, .semi-modal, [role="dialog"], [class*="drawer"], [class*="popover"]'))
        .filter(isVisible)
        .filter((el) => {
          const rect = el.getBoundingClientRect();
          return rect.left < window.innerWidth && rect.right > 0 && /邀约详情|博主微信/.test(clean(el.innerText || el.textContent || ''));
        })
        .sort((a, b) => b.getBoundingClientRect().right - a.getBoundingClientRect().right)[0] || document.body;
      const labels = Array.from(surface.querySelectorAll('*')).filter((el) => {
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0 && clean(el.innerText || el.textContent || '') === '博主微信';
      });
      for (const label of labels) {
        let row = label;
        while (row && row !== surface) {
          const text = clean(row.innerText || row.textContent || '');
          if (text.includes('博主微信') && row.querySelector('.item-right, [class*="right"]')) break;
          row = row.parentElement;
        }
        const right = row?.querySelector('.item-right, [class*="right"]');
        const local = clean(right?.innerText || right?.textContent || '');
        const firstLine = local.split(/\\s+/).find(Boolean);
        const value = valid(firstLine || local);
        if (value) return value;
      }
      const text = clean(document.body?.innerText || '');
      const labelMatch = text.match(/(?:博主微信|联系电话)\\s*[:：]?\\s*([^\\s|，,;；]+)/);
      if (labelMatch) return valid(labelMatch[1]);
      const nodes = Array.from(document.querySelectorAll('*')).filter((el) => {
        const rect = el.getBoundingClientRect();
        const nodeText = clean(el.innerText || el.textContent || '');
        return rect.width > 0 && rect.height > 0 && nodeText.includes('博主微信');
      });
      for (const node of nodes) {
        const local = clean(node.innerText || node.textContent || '');
        const localMatch = local.match(/(?:博主微信|联系电话)\\s*[:：]?\\s*([^\\s|，,;；]+)/);
        if (localMatch) {
          const value = valid(localMatch[1]);
          if (value) return value;
        }
        const parent = node.parentElement;
        const siblings = parent ? Array.from(parent.children).map((el) => clean(el.innerText || el.textContent || '')).filter(Boolean) : [];
        const idx = siblings.findIndex((value) => value.includes('博主微信'));
        if (idx >= 0 && siblings[idx + 1]) {
          const value = valid(siblings[idx + 1].replace(/^[:：]/, '').trim());
          if (value) return value;
        }
      }
      return '';
    })()
  `);
    return cleanText(contact);
}

async function waitForInviteDrawer(page) {
    await page.evaluate(`
    new Promise((resolve) => {
      const ready = () => {
        const drawers = Array.from(document.querySelectorAll('.d-drawer, .ant-drawer, [class*="drawer"], [role="dialog"]'));
        const drawer = drawers
          .filter((el) => {
            const rect = el.getBoundingClientRect();
            const style = getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';
          })
          .filter((el) => {
            const rect = el.getBoundingClientRect();
            return rect.left < window.innerWidth && rect.right > 0;
          })
          .sort((a, b) => b.getBoundingClientRect().right - a.getBoundingClientRect().right)[0];
        const text = drawer?.innerText || '';
        if (!/邀约详情/.test(text) || !/博主微信/.test(text)) return false;
        const rect = drawer.getBoundingClientRect();
        return rect.width > 0 && rect.left < window.innerWidth && rect.right > 0;
      };
      if (ready()) return resolve(true);
      if (!document.body) return resolve(false);
      const observer = new MutationObserver(() => {
        if (ready()) {
          observer.disconnect();
          resolve(true);
        }
      });
      observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      setTimeout(() => {
        observer.disconnect();
        resolve(false);
      }, 10000);
    })
  `);
}

async function readContactInfoFromInviteDrawer(page, invite, delayBounds) {
    const opened = await clickInviteDetailButtonForInvite(page, invite);
    if (!opened)
        return '';
    await humanPause(page, delayBounds, 0.8);
    await waitForInviteDrawer(page);
    await humanPause(page, delayBounds, 0.4);
    let contact = '';
    for (let attempt = 0; attempt < 3 && !contact; attempt++) {
        await clickRevealContactButton(page);
        await humanPause(page, delayBounds, 0.8 + attempt * 0.4);
        contact = await readWechatContact(page);
    }
    await closeDetailSurface(page);
    await humanPause(page, delayBounds, 0.5);
    return contact;
}

async function closeDetailSurface(page) {
    const closed = await clickText(page, ['关闭', '取消'], { exact: true });
    if (closed)
        return true;
    return page.evaluate(`
    (() => {
      const selectors = [
        '.ant-modal-close',
        '.semi-modal-close',
        '[aria-label="Close"]',
        '[class*="close"]'
      ];
      for (const selector of selectors) {
        const el = document.querySelector(selector);
        if (el) {
          el.click();
          return true;
        }
      }
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keyup', { key: 'Escape', bubbles: true }));
      return false;
    })()
  `);
}

async function clickInviteProfile(page, rowIndex) {
    return clickElementByEval(page, `
    () => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const isVisible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const rows = Array.from(document.querySelectorAll('${rowSelector()}'))
        .filter((el) => {
          const text = clean(el.innerText || el.textContent || '');
          return isVisible(el) && text.includes('查看详情') && text.includes('感兴趣') && !text.includes('不感兴趣');
        });
      const row = rows[${Number(rowIndex)}];
      if (!row) return null;
      const blocked = (el) => /查看详情|查看|感兴趣|不感兴趣|已拒绝|待确认/.test(clean(el.innerText || el.textContent || ''));
      const selectors = [
        '.d-avatar',
        '.kol-name',
        'a[href*="pgy.xiaohongshu.com"][href*="/solar/"]',
        'a[href*="/solar/"]',
        'a[href*="/kol"]',
        'a[href*="/profile"]',
        '[class*="avatar"]',
        '[class*="nick"]',
        '[class*="name"]',
        '[class*="author"]',
        '[class*="blogger"]'
      ];
      for (const selector of selectors) {
        const preferred = Array.from(row.querySelectorAll(selector)).find((el) => isVisible(el) && !blocked(el));
        if (preferred) return preferred.closest('a, button, [role="button"]') || preferred;
      }
      const cells = Array.from(row.querySelectorAll('td, [class*="cell"], [class*="column"]')).filter(isVisible);
      const infoCell = cells.find((cell) => !blocked(cell) && clean(cell.innerText || cell.textContent || '').length > 0);
      return infoCell || null;
    }
  `);
}

async function waitForPgyDetail(page, previousUrl) {
    await page.wait({ time: DETAIL_SETTLE_SECONDS });
    const ok = await page.evaluate(`
    (() => {
      const text = document.body?.innerText || '';
      const previousUrl = ${JSON.stringify(previousUrl || '')};
      const href = location.href;
      const isDetailUrl = /\\/solar\\/pre-trade\\/blogger-detail\\//.test(href);
      const hasProfileFields = /笔记主页/.test(text) && /小红书号|达人ID|达人昵称/.test(text);
      return href !== previousUrl && (isDetailUrl || hasProfileFields);
    })()
  `);
    if (!ok) {
        throw new EmptyResultError('xiaohongshu/pgyexcel', 'Creator Pugongying detail page did not load expected profile fields');
    }
}

async function readCreatorProfile(page, contact, source = {}) {
    const domData = await page.evaluate(`
    (() => {
      const clean = (value) => (value || '').replace(/\\s+/g, ' ').trim();
      const visible = (el) => {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const allText = clean(document.body?.innerText || '');
      const profile = document.querySelector('.blogger-base-info') || document.querySelector('[class*="blogger-base-info"]');
      const blocks = Array.from(document.querySelectorAll('*'))
        .filter((el) => visible(el) && clean(el.innerText || el.textContent || '').includes('笔记主页'))
        .map((el) => ({ el, text: clean(el.innerText || el.textContent || ''), rect: el.getBoundingClientRect() }))
        .filter((item) => item.text.length < 2000)
        .sort((a, b) => {
          const score = (item) =>
            (item.text.includes('小红书号') ? -1000 : 0)
            + (item.text.includes('达人昵称') ? -300 : 0)
            + item.rect.top
            + item.rect.left / 10
            + item.text.length / 100;
          return score(a) - score(b);
        });
      const block = blocks[0]?.el;
      const noteHomeText = clean(profile?.innerText || profile?.textContent || block?.innerText || block?.textContent || allText);
      const labels = Array.from(document.querySelectorAll('.blogger-tag-list .d-tag, .blogger-tag-list [class*="tag"], [class*="blogger-tag"] .d-tag'))
        .filter(visible)
        .map((el) => clean(el.innerText || el.textContent || ''))
        .filter((text) => text && text.length <= 30 && !/笔记主页|直播主页|小红书号|达人ID|达人昵称|主页|查看|收藏|邀约|粉丝|获赞/.test(text));
      const links = Array.from((profile || block || document).querySelectorAll('a[href]')).map((a) => a.href).filter(Boolean);
      const homepageLink = links.find((href) => /xiaohongshu\\.com\\/user\\/profile\\//.test(href)) || '';
      const nicknameNode = (profile || block || document).querySelector('.blogger-name, [class*="nick"], [class*="name"], [class*="author"]');
      const idMatch = noteHomeText.match(/(?:小红书号|达人ID)\\s*[:：]?\\s*([A-Za-z0-9._-]+)/);
      return {
        contact: ${JSON.stringify(contact)},
        fallbackNickname: ${JSON.stringify(source.nickname || '')},
        fallbackHomepageUrl: ${JSON.stringify(source.homepageUrl || '')},
        pgyUrl: location.href,
        homepageUrl: homepageLink,
        nickname: clean(nicknameNode?.innerText || nicknameNode?.textContent || ''),
        creatorId: idMatch?.[1] || '',
        labels,
        links,
        noteHomeText,
        bodyText: allText,
      };
    })()
  `);
    return parseCreatorProfileFromDomData(domData);
}

async function goBackToList(page, delayBounds) {
    await page.evaluate(`() => history.back()`);
    await humanPause(page, delayBounds, 1.5);
    await ensurePgyReady(page);
}

async function goToNextPage(page, delayBounds) {
    const clicked = await clickText(page, ['下一页', '下一页 >', '>'], { exact: false });
    if (!clicked)
        return false;
    await humanPause(page, delayBounds, 1.4);
    return true;
}

async function collectCurrentPage(page, delayBounds, pageNum) {
    const inviteRows = await fetchInterestedInvites(page, pageNum);
    const results = [];
    for (let i = 0; i < inviteRows.length; i++) {
        await humanPause(page, delayBounds, 0.7);
        const invite = { ...inviteRows[i], rowIndex: i };
        let contact = await readContactInfoFromInviteDrawer(page, invite, delayBounds);
        if (!contact)
            contact = await readContactInfoByKolId(page, invite.kolId);
        await humanPause(page, delayBounds, 0.5);
        await openCreatorDetailByKolId(page, invite.kolId, delayBounds);
        const creatorRow = await readCreatorProfile(page, contact, {
            nickname: invite.kolName,
            homepageUrl: `https://www.xiaohongshu.com/user/profile/${invite.kolId}`,
        });
        if (creatorRow['达人昵称'] || creatorRow['达人ID'] || creatorRow['蒲公英连接']) {
            results.push(creatorRow);
        }
        await goBackToList(page, delayBounds);
        await applyInterestedInviteFilter(page, delayBounds);
        await waitForInviteTable(page);
        await gotoListPageNumber(page, pageNum, delayBounds);
    }
    return results;
}

export async function runPgyExcelExport(page, kwargs) {
    const { startPage, endPage } = normalizePageRange(kwargs);
    const delayBounds = clampDelayBounds(kwargs['min-delay'], kwargs['max-delay']);
    const output = String(kwargs.output || './xiaohongshu-pgy-creators.csv');
    const rows = [];

    await page.goto(PGY_INVITE_LIST_URL);
    await humanPause(page, delayBounds, 1.2);
    await ensurePgyReady(page);
    await applyInterestedInviteFilter(page, delayBounds);
    await waitForInviteTable(page);
    await gotoListPageNumber(page, startPage, delayBounds);

    for (let currentPage = startPage; currentPage <= endPage; currentPage++) {
        const pageRows = await collectCurrentPage(page, delayBounds, currentPage);
        rows.push(...pageRows);
        const outputPath = writeCsv(output, rows);
        if (currentPage === endPage) {
            return {
                rows,
                outputPath,
                startPage,
                endPage,
                debug: rows.length === 0 ? await readInviteTableDebug(page) : undefined,
            };
        }
        const advanced = await goToNextPage(page, delayBounds);
        if (!advanced)
            break;
    }

    return {
        rows,
        outputPath: writeCsv(output, rows),
        startPage,
        endPage,
        debug: await readInviteTableDebug(page),
    };
}

cli({
    site: 'xiaohongshu',
    name: 'pgyexcel',
    description: '导出蒲公英感兴趣邀约博主资料为 Excel 可打开的 CSV',
    domain: 'pgy.xiaohongshu.com',
    strategy: Strategy.COOKIE,
    browser: true,
    navigateBefore: false,
    timeoutSeconds: 900,
    args: [
        { name: 'start-page', type: 'int', default: 1, help: '起始页，默认 1' },
        { name: 'end-page', type: 'int', help: '结束页，默认等于起始页' },
        { name: 'output', default: './xiaohongshu-pgy-creators.csv', help: 'CSV 输出路径' },
        { name: 'min-delay', type: 'number', default: 1.2, help: '每步最小等待秒数' },
        { name: 'max-delay', type: 'number', default: 3.8, help: '每步最大等待秒数' },
    ],
    columns: ['status', 'rows', 'output', 'pages'],
    validateArgs: (kwargs) => {
        normalizePageRange(kwargs);
        clampDelayBounds(kwargs['min-delay'], kwargs['max-delay']);
    },
    func: async (page, kwargs) => {
        const result = await runPgyExcelExport(page, kwargs);
        if (result.rows.length === 0) {
            const sample = result.debug?.samples?.join(' || ') || 'no visible table row sample';
            throw new EmptyResultError('xiaohongshu/pgyexcel', `No interested invite creators were exported from the selected pages. Table sample: ${sample}`);
        }
        return [{
            status: 'exported',
            rows: result.rows.length,
            output: result.outputPath,
            pages: `${result.startPage}-${result.endPage}`,
        }];
    },
});
