import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { describe, expect, it } from 'vitest';
import { getRegistry } from '@jackwener/opencli/registry';
import {
    CSV_COLUMNS,
    escapeCsvCell,
    extractInviteRowsFromDomData,
    isInterestedInviteText,
    normalizePageRange,
    parseCreatorProfileFromDomData,
    rowsToCsv,
    writeCsv,
} from './pgyexcel.js';
import './pgyexcel.js';

describe('xiaohongshu pgyexcel', () => {
    it('registers the command with expected args and navigation hardening', () => {
        const cmd = getRegistry().get('xiaohongshu/pgyexcel');
        expect(cmd).toBeDefined();
        expect(cmd.navigateBefore).toBe(false);
        expect(cmd.columns).toEqual(['status', 'rows', 'output', 'pages']);
        expect(cmd.args.map((arg) => arg.name)).toEqual([
            'start-page',
            'end-page',
            'output',
            'min-delay',
            'max-delay',
        ]);
    });

    it('normalizes and validates page ranges', () => {
        expect(normalizePageRange({})).toEqual({ startPage: 1, endPage: 1 });
        expect(normalizePageRange({ 'start-page': 2, 'end-page': 4 })).toEqual({ startPage: 2, endPage: 4 });
        expect(() => normalizePageRange({ 'start-page': 0 })).toThrow('--start-page');
        expect(() => normalizePageRange({ 'start-page': 3, 'end-page': 2 })).toThrow('--end-page');
    });

    it('escapes CSV cells and writes a UTF-8 BOM for Excel', () => {
        expect(escapeCsvCell('plain')).toBe('plain');
        expect(escapeCsvCell('a,b')).toBe('"a,b"');
        expect(escapeCsvCell('a"b')).toBe('"a""b"');
        expect(escapeCsvCell('a\nb')).toBe('"a\nb"');

        const csv = rowsToCsv([{
            '联系方式': 'wx,abc',
            '达人昵称': '昵称"一"',
            '达人ID': 'xhs001',
            '达人类型1': '美妆',
            '达人类型2': '',
            '蒲公英连接': 'https://pgy.example/detail',
            '主页链接': 'https://www.xiaohongshu.com/user/profile/abc',
        }]);
        expect(csv.startsWith('\uFEFF')).toBe(true);
        expect(csv).toContain(CSV_COLUMNS.join(','));
        expect(csv).toContain('"wx,abc"');
        expect(csv).toContain('"昵称""一"""');
    });

    it('writes CSV files to nested output paths', () => {
        const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'opencli-pgyexcel-'));
        const output = path.join(tempDir, 'nested', 'pgy.csv');
        const resolved = writeCsv(output, [{
            '联系方式': 'wx',
            '达人昵称': '测试达人',
            '达人ID': 'xhs123',
            '达人类型1': '穿搭',
            '达人类型2': '母婴',
            '蒲公英连接': 'https://pgy.xiaohongshu.com/a',
            '主页链接': 'https://www.xiaohongshu.com/user/profile/u',
        }]);
        expect(resolved).toBe(path.resolve(output));
        expect(fs.readFileSync(output, 'utf-8')).toContain('测试达人');
    });

    it('identifies interested invite text without matching not interested rows', () => {
        expect(isInterestedInviteText('博主意向 感兴趣 查看详情')).toBe(true);
        expect(isInterestedInviteText('博主意向 不感兴趣 查看详情')).toBe(false);
        expect(isInterestedInviteText('博主意向 待确认 查看详情')).toBe(false);
    });

    it('parses creator profile data with two, one, or no labels', () => {
        expect(parseCreatorProfileFromDomData({
            contact: 'wx001',
            pgyUrl: 'https://pgy.xiaohongshu.com/detail/1',
            homepageUrl: 'https://www.xiaohongshu.com/user/profile/u1',
            nickname: '达人一',
            creatorId: 'red001',
            labels: ['美妆', '护肤'],
        })).toEqual({
            '联系方式': 'wx001',
            '达人昵称': '达人一',
            '达人ID': 'red001',
            '达人类型1': '美妆',
            '达人类型2': '护肤',
            '蒲公英连接': 'https://pgy.xiaohongshu.com/detail/1',
            '主页链接': 'https://www.xiaohongshu.com/user/profile/u1',
        });

        expect(parseCreatorProfileFromDomData({
            noteHomeText: '笔记主页 达人昵称: 达人二 小红书号: red002',
            labels: ['旅行'],
            links: ['https://www.xiaohongshu.com/user/profile/u2'],
            pgyUrl: 'https://pgy.xiaohongshu.com/detail/2',
        })).toMatchObject({
            '达人昵称': '达人二',
            '达人ID': 'red002',
            '达人类型1': '旅行',
            '达人类型2': '',
            '主页链接': 'https://www.xiaohongshu.com/user/profile/u2',
        });

        expect(parseCreatorProfileFromDomData({
            noteHomeText: '笔记主页 小红书号: red003',
            labels: [],
        })).toMatchObject({
            '达人ID': 'red003',
            '达人类型1': '',
            '达人类型2': '',
        });
    });

    it('extracts only interested list rows and tolerates missing contact/detail fields', () => {
        expect(extractInviteRowsFromDomData([
            { nickname: '达人一', detailText: '达人一 感兴趣 查看详情', hasDetailButton: true, hasProfileTarget: true },
            { nickname: '达人二', detailText: '达人二 不感兴趣 查看详情', hasDetailButton: true, hasProfileTarget: true },
            { nickname: '达人三', detailText: '达人三 待确认 查看详情', hasDetailButton: true, hasProfileTarget: true },
            { nickname: '', detailText: '', hasDetailButton: false, hasProfileTarget: false },
        ])).toEqual([
            {
                index: 0,
                nickname: '达人一',
                detailText: '达人一 感兴趣 查看详情',
                hasDetailButton: true,
                hasProfileTarget: true,
            },
        ]);
    });
});
