import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";

const html = readFileSync(new URL("./index.html", import.meta.url), "utf8");
const script = html.match(/<script>([\s\S]*?)<\/script>/)?.[1];
const katexScript = html.match(/<script id="katexOfflineRuntime">([\s\S]*?)<\/script>/)?.[1];

assert.ok(script, "index.html should contain an inline script");
assert.ok(katexScript, "index.html should contain the offline KaTeX runtime");

const sandbox = { window: {}, Blob, DataView, TextEncoder, Uint8Array };
vm.createContext(sandbox);
vm.runInContext(script, sandbox);
vm.runInContext(katexScript, sandbox);
sandbox.window.katex = sandbox.katex;

const { buildClipboardHtml, buildDocxBlob, createStoredZip, markdownToHtml, mathMlNodeToOmml } = sandbox.window.md2doc;

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

test("renders mermaid code fences as diagram placeholders", () => {
  const source = [
    "```mermaid",
    "flowchart LR",
    "  A[开始] --> B[结束]",
    "```",
  ].join("\n");

  const rendered = markdownToHtml(source);

  assert.match(rendered, /class="mermaid-diagram"/);
  assert.match(rendered, /data-mermaid="/);
  assert.match(rendered, /flowchart LR/);
  assert.match(rendered, /A\[开始\] --&gt; B\[结束\]/);
  assert.doesNotMatch(rendered, /language-mermaid/);
});

test("renders inline and display math placeholders", () => {
  const source = [
    "行内公式：$E = mc^2$",
    "",
    "\\[",
    "P(X=k)=\\binom{n}{k}p^k(1-p)^{n-k}",
    "\\]",
  ].join("\n");

  const rendered = markdownToHtml(source);

  assert.match(rendered, /class="math-expression math-inline/);
  assert.match(rendered, /class="math-expression math-display/);
  assert.match(rendered, /data-display="false"/);
  assert.match(rendered, /data-display="true"/);
  assert.match(rendered, /<math/);
  assert.match(rendered, /<msup>/);
  assert.match(rendered, /<mfrac/);
});

test("accepts copied display math with bare square bracket delimiters", () => {
  const rendered = markdownToHtml([
    "[",
    "10\\binom92",
    "]",
  ].join("\n"));

  assert.match(rendered, /class="math-expression math-display/);
  assert.match(rendered, /data-display="true"/);
  assert.match(rendered, /<mfrac/);
  assert.doesNotMatch(rendered, /<p>\[<\/p>/);
});

test("renders boxed Chinese text with a visible border", () => {
  const rendered = markdownToHtml([
    "[",
    "\\boxed{\\text{古典概率的前提是基本结果等可能}}",
    "]",
  ].join("\n"));

  assert.match(rendered, /<menclose[^>]*notation="box"/);
  assert.match(rendered, /<menclose[^>]*style="[^"]*border:1px solid currentColor/);
  assert.match(rendered, /古典概率的前提是基本结果等可能/);
});

test("treats an unescaped percent sign as a literal percentage in formulas", () => {
  const rendered = markdownToHtml([
    "[",
    "\\boxed{50.7%}",
    "]",
  ].join("\n"));
  const encodedLatex = rendered.match(/data-latex="([^"]+)"/)?.[1] || "";

  assert.match(rendered, /<menclose[^>]*notation="box"/);
  assert.match(rendered, /<menclose[^>]*style="[^"]*display:inline-block/);
  assert.doesNotMatch(rendered, /math-error/);
  assert.equal(decodeURIComponent(encodedLatex), "\\boxed{50.7\\%}");

  const alreadyEscaped = markdownToHtml("$50.7\\%$");
  const escapedLatex = alreadyEscaped.match(/data-latex="([^"]+)"/)?.[1] || "";
  assert.equal(decodeURIComponent(escapedLatex), "50.7\\%");
});

test("normalizes copied multi-line equation chains", () => {
  const first = markdownToHtml([
    "[",
    "P(\\text{都是女性})",
    "==============",
    "",
    "# \\frac{10}{45}",
    "",
    "\\frac29",
    "]",
  ].join("\n"));
  const firstLatex = decodeURIComponent(first.match(/data-latex="([^"]+)"/)?.[1] || "");

  assert.equal(firstLatex, "P(\\text{都是女性}) = \\frac{10}{45} = \\frac29");
  assert.doesNotMatch(first, /math-error/);
  assert.match(first, /都是女性/);

  const second = markdownToHtml([
    "[",
    "\\binom21\\binom82",
    "================",
    "",
    "# 2\\times28",
    "",
    "56",
    "]",
  ].join("\n"));
  const secondLatex = decodeURIComponent(second.match(/data-latex="([^"]+)"/)?.[1] || "");

  assert.equal(secondLatex, "\\binom21\\binom82 = 2\\times28 = 56");
  assert.doesNotMatch(second, /math-error/);
  assert.match(second, /<mfrac/g);
});

test("does not treat regular Markdown links as display math", () => {
  const rendered = markdownToHtml("[打开项目](https://github.com/czhcc/md2docx)");

  assert.match(rendered, /<a href=/);
  assert.doesNotMatch(rendered, /math-display/);
});

test("converts MathML structures to editable OMML", () => {
  const element = (localName, children = [], textContent = "", attributes = {}) => ({
    nodeType: 1,
    localName,
    children,
    childNodes: children,
    textContent,
    getAttribute(name) {
      return attributes[name] || "";
    },
  });
  const numerator = element("mi", [], "n");
  const denominator = element("mn", [], "2");
  const fraction = element("mfrac", [numerator, denominator]);
  const power = element("msup", [element("mi", [], "x"), element("mn", [], "2")]);
  const boxed = element("menclose", [element("mtext", [], "结论")], "结论", { notation: "box" });
  const math = element("math", [fraction, power, boxed]);

  const omml = mathMlNodeToOmml(math);

  assert.match(omml, /<m:f>/);
  assert.match(omml, /<m:num>/);
  assert.match(omml, /<m:den>/);
  assert.match(omml, /<m:sSup>/);
  assert.match(omml, /<m:t[^>]*>x<\/m:t>/);
  assert.match(omml, /<m:borderBox>/);
  assert.match(omml, /<m:t[^>]*>结论<\/m:t>/);
});

test("creates a DOCX-compatible ZIP containing editable equation XML", async () => {
  const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p></w:body></w:document>';
  const blob = createStoredZip([{ name: "word/document.xml", data: documentXml }]);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const decoded = new TextDecoder().decode(bytes);

  assert.equal(bytes[0], 0x50);
  assert.equal(bytes[1], 0x4b);
  assert.match(decoded, /word\/document\.xml/);
  assert.match(decoded, /<m:oMath>/);
});

test("builds a complete DOCX package with an editable inline formula", async () => {
  const mathToken = (localName, textContent = "", children = []) => ({
    nodeType: 1,
    localName,
    children,
    childNodes: children,
    textContent,
    getAttribute() {
      return "";
    },
  });
  const math = mathToken("math", "", [
    mathToken("msup", "", [mathToken("mi", "x"), mathToken("mn", "2")]),
  ]);
  const formula = {
    nodeType: 1,
    tagName: "SPAN",
    classList: { contains: (name) => name === "math-expression" },
    dataset: { display: "false", latex: encodeURIComponent("x^2") },
    childNodes: [],
    querySelector: (selector) => selector === "math" ? math : null,
  };
  const paragraph = {
    nodeType: 1,
    tagName: "P",
    classList: { contains: () => false },
    childNodes: [{ nodeType: 3, nodeValue: "公式：" }, formula],
  };
  const preview = {
    children: [paragraph],
    querySelectorAll: () => [],
  };

  const blob = await buildDocxBlob(preview);
  const decoded = new TextDecoder().decode(new Uint8Array(await blob.arrayBuffer()));

  assert.match(decoded, /\[Content_Types\]\.xml/);
  assert.match(decoded, /word\/document\.xml/);
  assert.match(decoded, /word\/styles\.xml/);
  assert.match(decoded, /<m:oMath>/);
  assert.match(decoded, /<m:sSup>/);
});

test("page embeds Mermaid for offline use and exposes a render hook", () => {
  assert.match(html, /id="mermaidOfflineRuntime"/);
  assert.match(html, /globalThis\["mermaid"\]/);
  assert.doesNotMatch(html, /cdn\.jsdelivr\.net\/npm\/mermaid/);
  assert.doesNotMatch(html, /import mermaid from/);
  assert.match(html, /mermaid\.initialize\(\{[\s\S]*?startOnLoad: false/);
  assert.match(html, /renderMermaidDiagrams/);
  assert.match(html, /window\.md2docRenderMermaid/);
  assert.match(html, /prepareMermaidImagesForCopy/);
  assert.match(html, /toDataURL\("image\/png"\)/);
});

test("page embeds KaTeX and exports editable OMML equations to DOCX", () => {
  assert.match(html, /id="katexOfflineRuntime"/);
  assert.match(html, /id="exportDocxButton"/);
  assert.match(html, /output: "mathml"/);
  assert.match(html, /function mathMlNodeToOmml/);
  assert.match(html, /<m:oMathPara>/);
  assert.match(html, /<m:oMath>/);
  assert.match(html, /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document/);
  assert.match(html, /function buildDocxBlob/);
  assert.match(html, /function exportDocx/);
  assert.doesNotMatch(html, /cdn\.jsdelivr\.net\/npm\/katex/);
});

test("rich clipboard prepares formula images for stable Word appearance", () => {
  assert.match(html, /prepareMathImagesForCopy/);
  assert.match(html, /mathElementToPngDataUrl/);
  assert.match(html, /data-md2doc-math-image="true"/);
  assert.match(html, /replaceMathExpressionsInClone/);
});

test("page has no external executable or stylesheet dependencies", () => {
  assert.doesNotMatch(html, /<script[^>]+src=["']https?:\/\//i);
  assert.doesNotMatch(html, /<link[^>]+href=["']https?:\/\//i);
  assert.doesNotMatch(html, /import\s+[\s\S]*?from\s+["']https?:\/\//i);
});

test("page exposes rich clipboard html copy support", () => {
  assert.match(html, /ClipboardItem/);
  assert.match(html, /text\/html/);
  assert.match(html, /text\/plain/);
  assert.match(html, /copyRichText/);
});

test("default markdown includes a local image copy sample", () => {
  assert.match(html, /!\[本地图片复制测试\]\(images\.png\)/);
});

test("clipboard html constrains regular images to the Word page width", () => {
  const preview = {
    innerHTML: '<p><img src="images.png" alt="本地图片复制测试"></p>',
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<img style="[^"]*max-width: 100%[^"]*" src="images\.png"/);
  assert.match(clipboardHtml, /height: auto/);
});

test("page exposes a floating back-to-top button when copy button scrolls away", () => {
  assert.match(html, /id="backToTopButton"/);
  assert.match(html, /class="back-to-top"/);
  assert.match(html, /updateBackToTopVisibility/);
  assert.match(html, /IntersectionObserver/);
  assert.match(html, /scrollTo\(\{ top: 0, behavior: "smooth" \}\)/);
});

test("page exposes a collapsible Markdown input drawer", () => {
  assert.match(html, /id="workspace"/);
  assert.match(html, /id="sourcePane"/);
  assert.match(html, /id="sourceDrawerToggle"/);
  assert.match(html, /aria-controls="sourcePane"/);
  assert.match(html, /grid-template-columns: 44px minmax\(320px, 1fr\)/);
  assert.match(html, /classList\.toggle\("source-collapsed", isCollapsed\)/);
  assert.match(html, /setSourceDrawerCollapsed/);
  assert.match(html, /setupSourceDrawer/);
});

test("page can open a local Markdown file into the input", () => {
  assert.match(html, /id="openMarkdownButton"/);
  assert.match(html, /id="markdownFileInput"/);
  assert.match(html, /accept="\.md,\.markdown,text\/markdown,text\/plain"/);
  assert.match(html, /await file\.text\(\)/);
  assert.match(html, /input\.value = content/);
  assert.match(html, /setupMarkdownFileOpen/);
  assert.match(html, /fileInput\.value = ""/);
});

test("paste button replaces Markdown only after clipboard text is read", async () => {
  const previousDocument = sandbox.document;
  const previousNavigator = sandbox.navigator;
  const previousSetTimeout = sandbox.window.setTimeout;
  const input = {
    value: "原有内容",
    focus() {},
    setSelectionRange() {},
  };
  const preview = { innerHTML: "" };
  const sourceCount = { textContent: "" };
  const button = {
    textContent: "粘贴",
    title: "",
    disabled: false,
    classList: { add() {}, remove() {} },
    addEventListener(event, handler) {
      if (event === "click") this.clickHandler = handler;
    },
  };
  const elements = { pasteMarkdownButton: button, markdownInput: input, preview, sourceCount };

  try {
    sandbox.document = { getElementById: (id) => elements[id] || null };
    sandbox.window.setTimeout = () => 0;
    sandbox.navigator = { clipboard: { readText: async () => "# 新内容" } };
    sandbox.window.md2doc.setupMarkdownPaste();

    await button.clickHandler();
    assert.equal(input.value, "# 新内容");
    assert.match(preview.innerHTML, /<h1>新内容<\/h1>/);
    assert.equal(sourceCount.textContent, "5 字");

    sandbox.navigator.clipboard.readText = async () => { throw new Error("denied"); };
    await button.clickHandler();
    assert.equal(input.value, "# 新内容");
    assert.equal(button.textContent, "读取失败");
  } finally {
    sandbox.document = previousDocument;
    sandbox.navigator = previousNavigator;
    sandbox.window.setTimeout = previousSetTimeout;
  }
});

test("page exposes an interactive Mermaid zoom modal", () => {
  assert.match(html, /id="mermaidModal"/);
  assert.match(html, /role="dialog"/);
  assert.match(html, /id="mermaidModalViewport"/);
  assert.match(html, /MERMAID_MODAL_DEFAULT_SCALE = 1\.75/);
  assert.match(html, /openMermaidModal/);
  assert.match(html, /closeMermaidModal/);
  assert.match(html, /setMermaidModalZoom/);
  assert.match(html, /addEventListener\("wheel"/);
  assert.match(html, /addEventListener\("pointermove"/);
  assert.match(html, /event\.key === "Escape"/);
  assert.match(html, /fontSize: "17px"/);
  assert.match(html, /setProperty\("max-width", "none", "important"\)/);
});

test("page exposes a floating h1 and h2 preview table of contents", () => {
  assert.match(html, /id="previewShell"/);
  assert.match(html, /id="previewToc"/);
  assert.match(html, /id="previewTocList"/);
  assert.match(html, /position: fixed/);
  assert.match(html, /right: 24px/);
  assert.match(html, /querySelectorAll\("h1, h2"\)/);
  assert.match(html, /heading\.scrollIntoView\(\{ behavior: "smooth", block: "start" \}\)/);
  assert.match(html, /previewTocState\.lockedTargetId = button\.dataset\.tocTarget/);
  assert.match(html, /setPreviewTocActive\(previewTocState\.lockedTargetId\)/);
  assert.match(html, /previewShell\.scrollHeight > previewShell\.clientHeight \+ 1/);
  assert.match(html, /shellCanScroll && previewShell\.scrollTop \+ previewShell\.clientHeight/);
  assert.match(html, /!shellCanScroll && window\.scrollY \+ window\.innerHeight/);
  assert.match(html, /Math\.max\(0, shellRect\.top\) \+ 72/);
  assert.match(html, /setTimeout\(\(\) => \{[\s\S]*?updatePreviewTocActive\(\);[\s\S]*?\}, 180\)/);
  assert.match(html, /updatePreviewTocPosition/);
  assert.match(html, /window\.addEventListener\("scroll", scheduleActiveUpdate/);
  assert.match(html, /buildPreviewToc/);
  assert.match(html, /setupPreviewToc/);
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

  assert.match(clipboardHtml, /<div style="[^"]*width: 100%/);
  assert.match(clipboardHtml, /<div style="[^"]*max-width: 100%/);
  assert.match(clipboardHtml, /<table[^>]+width="100%"/);
  assert.match(clipboardHtml, /<table style="[^"]*width: 100%/);
  assert.match(clipboardHtml, /mso-width-percent: 1000/);
  assert.match(clipboardHtml, /<td style="[^"]*overflow-wrap: anywhere/);
  assert.doesNotMatch(clipboardHtml, /<table[^>]+width="600"/);
  assert.doesNotMatch(clipboardHtml, /<div style="[^"]*width: 600px/);
  assert.match(html, /container\.style\.width = WORD_PAGE_BODY_WIDTH/);
  assert.doesNotMatch(html, /container\.style\.width = "800px"/);
});

test("clipboard html fixes multi-column tables to the Word page width", () => {
  const preview = {
    innerHTML: [
      "<table>",
      "<thead><tr>",
      "<th>主题</th><th>标签</th><th>范围</th><th>说明</th><th>备注</th>",
      "</tr></thead>",
      "<tbody><tr>",
      "<td>现有主题整理后的标签</td>",
      "<td>特别长的标签内容会自动换行不会撑出页面</td>",
      "<td>范围</td><td>说明</td><td>备注</td>",
      "</tr></tbody>",
      "</table>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<table[^>]+width="100%"/);
  assert.match(clipboardHtml, /table-layout: fixed/);
  assert.match(clipboardHtml, /width: 100%/);
  assert.match(clipboardHtml, /mso-width-percent: 1000/);
  assert.match(clipboardHtml, /mso-table-layout-alt: fixed/);
  assert.match(clipboardHtml, /<colgroup><col style="width: 20%;">/);
  assert.match(clipboardHtml, /word-break: break-all/);
  assert.doesNotMatch(clipboardHtml, /<table[^>]*style="[^"]*width: 600px/);
});

test("clipboard html adds soft break points to long table field names", () => {
  const preview = {
    innerHTML: [
      "<table>",
      "<thead><tr>",
      "<th>sort</th>",
      "<th>label_code</th>",
      "<th>标签名称</th>",
      "<th>label_group</th>",
      "<th>label_type</th>",
      "<th>resource_scope</th>",
      "</tr></thead>",
      "<tbody><tr>",
      "<td>100</td>",
      "<td>ACTIVE_THEME</td>",
      "<td>现有主题整理后的标签</td>",
      "<td>NATIONAL_SOVEREIGNTY</td>",
      "<td>topic_label</td>",
      "<td>video_frame/text_resource</td>",
      "</tr></tbody>",
      "</table>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<table[^>]+width="100%"/);
  assert.match(clipboardHtml, /<colgroup><col style="width: 16\.6667%;">/);
  assert.match(clipboardHtml, /label_&#8203;code/);
  assert.match(clipboardHtml, /resource_&#8203;scope/);
  assert.match(clipboardHtml, /video_&#8203;frame\/&#8203;text_&#8203;resource/);
  assert.doesNotMatch(clipboardHtml, /<table[^>]*style="[^"]*width: 600px/);
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

test("clipboard html converts rendered Mermaid diagrams to pasted images", () => {
  const preview = {
    innerHTML: [
      '<div class="mermaid-diagram" data-word-image="data:image/png;base64,AAAA">',
      '<svg viewBox="0 0 100 50"><text>流程图</text></svg>',
      "</div>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /data-md2doc-mermaid-image="true"/);
  assert.match(clipboardHtml, /<img[^>]+src="data:image\/png;base64,AAAA"/);
  assert.match(clipboardHtml, /max-width: 100%/);
  assert.doesNotMatch(clipboardHtml, /<svg/);
});

test("clipboard html removes all Mermaid SVG label text after image conversion", () => {
  const preview = {
    innerHTML: [
      '<div class="mermaid-diagram" data-word-image="data:image/png;base64,BBBB">',
      '<svg viewBox="0 0 100 50">',
      '<foreignObject><div><span>输入 Markdown</span></div></foreignObject>',
      '<foreignObject><div><span>实时预览</span></div></foreignObject>',
      '<foreignObject><div><span>复制富文本</span></div></foreignObject>',
      "</svg>",
      "</div>",
    ].join(""),
  };

  const clipboardHtml = buildClipboardHtml(preview);

  assert.match(clipboardHtml, /<img[^>]+src="data:image\/png;base64,BBBB"/);
  assert.doesNotMatch(clipboardHtml, /输入 Markdown/);
  assert.doesNotMatch(clipboardHtml, /实时预览/);
  assert.doesNotMatch(clipboardHtml, /复制富文本/);
  assert.doesNotMatch(clipboardHtml, /foreignObject/);
});
