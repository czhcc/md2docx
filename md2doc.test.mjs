import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

const html = readFileSync(new URL("./index.html", import.meta.url), "utf8");
const script = html.match(/<script>([\s\S]*?)<\/script>/)?.[1];

assert.ok(script, "index.html should contain an inline script");

const sandbox = { window: {} };
vm.createContext(sandbox);
vm.runInContext(script, sandbox);

const { buildClipboardHtml, markdownToHtml } = sandbox.window.md2doc;

test("renders common markdown blocks for rich docx paste", () => {
  const source = [
    "# 标题",
    "",
    "这是一段 **重点** 和 `代码`。",
    "",
    "- 第一项",
    "- 第二项",
    "",
    "```js",
    "console.log('ok');",
    "```",
    "",
    "| 列 | 值 |",
    "| --- | --- |",
    "| A | B |",
  ].join("\n");

  const rendered = markdownToHtml(source);

  assert.match(rendered, /<h1>标题<\/h1>/);
  assert.match(rendered, /<strong>重点<\/strong>/);
  assert.match(rendered, /<code>代码<\/code>/);
  assert.match(rendered, /<ul>/);
  assert.match(rendered, /<li>第一项<\/li>/);
  assert.match(rendered, /<pre><code class="language-js">/);
  assert.match(rendered, /<table>/);
  assert.match(rendered, /<th>列<\/th>/);
  assert.match(rendered, /<td>B<\/td>/);
});

test("page exposes rich clipboard html copy support", () => {
  assert.match(html, /ClipboardItem/);
  assert.match(html, /text\/html/);
  assert.match(html, /text\/plain/);
  assert.match(html, /copyRichText/);
});

test("clipboard html carries Word-friendly inline table styling", () => {
  const preview = {
    innerHTML: [
      "<h2>2. 核心职责划分</h2>",
      "<table>",
      "<thead><tr><th>对象</th><th>作用</th></tr></thead>",
      "<tbody><tr><td>Task</td><td>任务与策略边界</td></tr></tbody>",
      "</table>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<h2 style="[^"]*border-bottom: 1px solid #d9e2ec/);
  assert.match(clipboardHtml, /<table style="[^"]*border-collapse: collapse/);
  assert.match(clipboardHtml, /<th style="[^"]*border: 1px solid #aab7c4/);
  assert.match(clipboardHtml, /<td style="[^"]*border: 1px solid #aab7c4/);
  assert.match(clipboardHtml, /mso-border-alt/);
});

test("clipboard output constrains tables to a Word page body width", () => {
  const preview = {
    innerHTML: "<table><tbody><tr><td>很长的表格内容</td><td>更多内容</td></tr></tbody></table>",
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /max-width: 600px/);
  assert.match(clipboardHtml, /<table style="[^"]*max-width: 100%/);
  assert.match(clipboardHtml, /<td style="[^"]*overflow-wrap: anywhere/);
  assert.match(html, /container\.style\.width = WORD_PAGE_BODY_WIDTH/);
  assert.doesNotMatch(html, /container\.style\.width = "800px"/);
});

test("clipboard html wraps fenced code in a light gray Word-friendly table cell", () => {
  const preview = {
    innerHTML: '<pre><code class="language-js">console.log(&#39;复制为富文本&#39;);</code></pre>',
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /data-md2doc-code-block="true"/);
  assert.match(clipboardHtml, /<td[^>]+bgcolor="#f3f4f6"/);
  assert.match(clipboardHtml, /mso-shading: #f3f4f6/);
  assert.match(clipboardHtml, /background-color: #f3f4f6/);
  assert.match(clipboardHtml, /color: #111827/);
  assert.match(clipboardHtml, /font-family: Consolas/);
  assert.match(clipboardHtml, /console\.log/);
});

test("clipboard html preserves section gaps, blockquote shading, and separators for Word", () => {
  const preview = {
    innerHTML: [
      "<h2>表格</h2>",
      "<table><tbody><tr><td>字段</td><td>说明</td></tr></tbody></table>",
      '<pre><code class="language-js">console.log(&#39;复制为富文本&#39;);</code></pre>',
      "<blockquote>点击上方按钮后，在 docx 编辑器中直接粘贴。</blockquote>",
      "<hr>",
      "<p>下一段</p>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<table data-md2doc-spacer="true"/);
  assert.match(clipboardHtml, /height: 14pt/);
  assert.match(clipboardHtml, /mso-height-rule: exactly/);
  assert.match(clipboardHtml, /border: none/);
  assert.match(clipboardHtml, /data-md2doc-quote="true"/);
  assert.match(clipboardHtml, /<td[^>]+bgcolor="#f3f7fc"/);
  assert.match(clipboardHtml, /mso-shading: #f3f7fc/);
  assert.match(clipboardHtml, /border-left: 3pt solid #1565c0/);
  assert.match(clipboardHtml, /data-md2doc-separator="true"/);
  assert.match(clipboardHtml, /border-top: 1px solid #cbd5e1/);
});
