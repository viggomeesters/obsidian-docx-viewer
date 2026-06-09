import fs from "node:fs";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = new URL("../fixtures/", import.meta.url);
fs.mkdirSync(root, { recursive: true });

function contentTypes(extraOverrides = "") {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  ${extraOverrides}
</Types>`;
}

function rels(items) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${items.join("\n")}
</Relationships>`;
}

function documentXml(body) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>${body}<w:sectPr/></w:body>
</w:document>`;
}

function paragraph(text, style = "", extra = "") {
  const styleXml = style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : "";
  return `<w:p>${styleXml}<w:r><w:t>${escapeXml(text)}</w:t></w:r>${extra}</w:p>`;
}

function listParagraph(text) {
  return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function hyperlink(text, rid) {
  return `<w:p><w:hyperlink r:id="${rid}"><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:hyperlink></w:p>`;
}

function trackedChanges() {
  return `<w:p><w:ins w:id="1"><w:r><w:t>Inserted review text</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>Deleted draft text</w:delText></w:r></w:del></w:p>`;
}

function table(rows) {
  return `<w:tbl>${rows.map((row) => `<w:tr>${row.map((cell) => `<w:tc><w:p><w:r><w:t>${escapeXml(cell)}</w:t></w:r></w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
}

function drawingParagraph() {
  return `<w:p><w:r><w:drawing><a:blip r:embed="rIdImage1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></w:drawing></w:r></w:p>`;
}

function commentsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2026-06-09T08:21:00Z"><w:p><w:r><w:t>Check the read-only scope.</w:t></w:r></w:p></w:comment>
</w:comments>`;
}

function notesXml(kind) {
  const tag = kind === "footnote" ? "footnote" : "endnote";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:${kind}s xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:${tag} w:id="1"><w:p><w:r><w:t>${kind === "footnote" ? "Footnote detail" : "Endnote detail"}</w:t></w:r></w:p></w:${tag}>
</w:${kind}s>`;
}

async function writeZip(name, build) {
  const zip = new JSZip();
  await build(zip);
  const buffer = await zip.generateAsync({ compression: "DEFLATE", type: "nodebuffer" });
  fs.writeFileSync(new URL(name, root), buffer);
}

async function addBaseDocument(zip) {
  zip.file("[Content_Types].xml", contentTypes());
  zip.file("_rels/.rels", rels([
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
  ]));
  zip.file("docProps/core.xml", '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>DOCX Viewer Fixture</dc:title><dc:creator>Codex</dc:creator></cp:coreProperties>');
  zip.file("docProps/app.xml", "<Properties><Application>Fixture Generator</Application></Properties>");
  zip.file("word/_rels/document.xml.rels", rels([
    '<Relationship Id="rIdHyperlink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/docx" TargetMode="External"/>',
    '<Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
    '<Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
    '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    '<Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
  ]));
  zip.file("word/document.xml", documentXml([
    paragraph("DOCX Viewer Fixture", "Title"),
    paragraph("Executive summary", "Heading1"),
    paragraph("This document proves text extraction."),
    listParagraph("First list item"),
    hyperlink("External reference", "rIdHyperlink1"),
    table([["Name", "Value"], ["Rows", "2"]]),
    drawingParagraph(),
    trackedChanges(),
  ].join("")));
  zip.file("word/comments.xml", commentsXml());
  zip.file("word/footnotes.xml", notesXml("footnote"));
  zip.file("word/endnotes.xml", notesXml("endnote"));
  zip.file("word/media/image1.png", Buffer.from("89504e470d0a1a0a0000000d49484452", "hex"));
}

await writeZip("simple.docx", addBaseDocument);
await writeZip("no-comments.docx", async (zip) => {
  await addBaseDocument(zip);
  zip.remove("word/comments.xml");
});
await writeZip("embedded-object.docx", async (zip) => {
  await addBaseDocument(zip);
  zip.file("word/embeddings/oleObject1.bin", "binary metadata only");
  zip.file("word/_rels/document.xml.rels", rels([
    '<Relationship Id="rIdOle1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>',
  ]));
});
await writeZip("large.docx", async (zip) => {
  zip.file("[Content_Types].xml", contentTypes());
  zip.file("word/document.xml", documentXml(Array.from({ length: 1100 }, (_, index) => paragraph(`Paragraph ${index + 1}`)).join("")));
});

fs.writeFileSync(new URL("malformed.docx", root), Buffer.from("not a zip"));
fs.writeFileSync(new URL("encrypted.docx", root), Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]));

function escapeXml(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

console.log(`DOCX fixtures written to ${fileURLToPath(root)}`);
