import { describe, expect, it } from 'vitest';
import type { OfficeApi } from '../office/client';
import { findGroupMembers, originalMessageLink } from '../gateway/mention-watch';

describe('mention backup metadata', () => {
  it('finds the owner and sender name across member pages', async () => {
    const calls: unknown[] = [];
    const office = { call: async (request: unknown) => {
      calls.push(request);
      return calls.length === 1
        ? { items: [{ member_id: 'ou_owner', name: '创建者' }], has_more: true, page_token: 'page-2' }
        : { items: [{ member_id: 'ou_sender', name: '张三' }], has_more: false };
    } } as unknown as OfficeApi;
    expect(await findGroupMembers(office, 'oc_group', 'ou_owner', 'ou_sender')).toEqual({ ownerPresent: true, senderName: '张三' });
    expect(calls).toHaveLength(2);
  });

  it('uses openChatId and message position for the group link', () => {
    const link = originalMessageLink('oc_group', '42');
    expect(link).toBe('https://applink.co.cibfintech.com/client/chat/open?openChatId=oc_group&position=42');
    expect(link).not.toContain('?chatId=');
  });

  it('opens the group without claiming message positioning when position is unavailable', () => {
    expect(originalMessageLink('oc_group')).toBe('https://applink.co.cibfintech.com/client/chat/open?openChatId=oc_group');
  });

  it('uses the confirmed intranet host and only accepts positioning for the same chat', () => {
    expect(originalMessageLink('oc_group', undefined, 'https://applink.feishu.cn/client/chat/open?openChatId=oc_group&position=7'))
      .toBe('https://applink.co.cibfintech.com/client/chat/open?openChatId=oc_group&position=7');
    expect(originalMessageLink('oc_group', undefined, 'https://applink.intra.example/client/chat/open?openChatId=oc_other&position=7'))
      .toBe('https://applink.co.cibfintech.com/client/chat/open?openChatId=oc_group');
  });
});
