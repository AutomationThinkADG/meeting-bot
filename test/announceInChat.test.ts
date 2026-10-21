// npx ts-node test/announceInChat.test.ts
// The bot's "<assistant> is transcribing this meeting." chat line, against
// small mock Teams pages in your installed Chrome. Real Teams markup changes;
// the live check is in vpm-api/docs/TESTING.md (Layer 2).
import assert from 'assert';
import { chromium } from 'playwright';
import { createLogger, transports } from 'winston';
import { announceInChat } from '../src/lib/meetingReadiness';

const logger = createLogger({ silent: true, transports: [new transports.Console()] });

const page = (body: string) => `<!doctype html><html><body>
  ${body}
  <ul id="sent"></ul>
  <script>
    const chat = document.querySelector('[data-tid="chat-button"], #chat-button');
    if (chat) chat.addEventListener('click', () => {
      const tpl = document.getElementById('compose');
      if (tpl) document.body.appendChild(tpl.content.cloneNode(true));
      const box = document.querySelector('[data-tid="ckeditor"]');
      if (box) box.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          const li = document.createElement('li');
          li.textContent = box.textContent;
          document.getElementById('sent').appendChild(li);
          box.textContent = '';
        }
      });
    });
  </script></body></html>`;

const COMPOSE = `<template id="compose"><div data-tid="ckeditor" contenteditable="true" role="textbox" aria-label="Type a message"></div></template>`;

(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  const ctx = await browser.newContext();
  const sent = async (p: import('playwright').Page) => p.$$eval('#sent li', (l) => l.map((x) => x.textContent));
  try {
    // 1. Normal meeting: one line, nothing else.
    let p = await ctx.newPage();
    await p.setContent(page(`<button data-tid="chat-button" aria-label="Chat">Chat</button>${COMPOSE}`));
    assert.strictEqual(await announceInChat(p, logger, 'ADG Transcript'), true);
    assert.deepStrictEqual(await sent(p), ['ADG Transcript is transcribing this meeting.']);

    // 2. The project's own assistant name; blank falls back to the default.
    p = await ctx.newPage();
    await p.setContent(page(`<button id="chat-button" aria-label="Show conversation">Chat</button>${COMPOSE}`));
    await announceInChat(p, logger, 'Nova');
    assert.deepStrictEqual(await sent(p), ['Nova is transcribing this meeting.']);
    p = await ctx.newPage();
    await p.setContent(page(`<button data-tid="chat-button">Chat</button>${COMPOSE}`));
    await announceInChat(p, logger, '  ');
    assert.deepStrictEqual(await sent(p), ['ADG Transcript is transcribing this meeting.']);

    // 3. No chat button (guests can't chat): returns false, never throws.
    p = await ctx.newPage();
    await p.setContent(page(`<button aria-label="People">People</button>`));
    assert.strictEqual(await announceInChat(p, logger, 'ADG Transcript'), false);

    // 4. Chat opens but there is no compose box (chat is read-only): false, no throw.
    p = await ctx.newPage();
    await p.setContent(page(`<button data-tid="chat-button">Chat</button>`));
    const t0 = Date.now();
    assert.strictEqual(await announceInChat(p, logger, 'ADG Transcript'), false);
    assert.ok(Date.now() - t0 < 15000, 'gives up quickly so recording starts');

    // 5. Switched off with TEAMS_ANNOUNCE_IN_CHAT=false.
    const config = require('../src/config').default;
    config.teamsAnnounceInChat = false;
    p = await ctx.newPage();
    await p.setContent(page(`<button data-tid="chat-button">Chat</button>${COMPOSE}`));
    assert.strictEqual(await announceInChat(p, logger, 'ADG Transcript'), false);
    assert.deepStrictEqual(await sent(p), []);

    console.log('ALL CHAT ANNOUNCEMENT TESTS PASSED');
  } finally {
    await browser.close();
  }
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
