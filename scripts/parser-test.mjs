import assert from "node:assert/strict";
import fs from "node:fs";
import esbuild from "esbuild";

await esbuild.build({
  bundle: true,
  entryPoints: ["src/parser.ts"],
  format: "esm",
  outfile: ".tmp-parser-test.mjs",
  platform: "node",
  target: "node20",
});

const {
  filterBlocks,
  filterComments,
  filterMedia,
  filterNotes,
  parseDocx,
} = await import(new URL("../.tmp-parser-test.mjs", import.meta.url));

function fixture(name) {
  const data = fs.readFileSync(new URL(`../fixtures/${name}`, import.meta.url));
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
}

const simple = await parseDocx(fixture("simple.docx"));
assert.equal(simple.summary.headingCount, 2);
assert.equal(simple.summary.tableCount, 1);
assert.equal(simple.summary.commentCount, 1);
assert.equal(simple.summary.noteCount, 2);
assert.equal(simple.summary.mediaCount, 1);
assert.equal(simple.summary.hyperlinkCount, 1);
assert.equal(simple.summary.trackedChangeCount, 2);
assert.equal(simple.title, "DOCX Viewer Fixture");
assert.ok(simple.blocks.some((block) => block.text.includes("Executive summary")));
assert.ok(simple.blocks.some((block) => block.rows.flat().includes("Rows")));
assert.ok(simple.comments[0].text.includes("read-only"));
assert.ok(simple.notes.some((note) => note.text.includes("Footnote detail")));
assert.ok(simple.media.some((item) => item.path === "word/media/image1.png"));
assert.equal(filterBlocks(simple.blocks, "executive").length, 1);
assert.equal(filterComments(simple.comments, "reviewer").length, 1);
assert.equal(filterNotes(simple.notes, "endnote").length, 1);
assert.equal(filterMedia(simple.media, ".png").length, 1);
assert.ok(simple.warnings.some((warning) => warning.includes("external relationships")));
assert.ok(simple.warnings.some((warning) => warning.includes("tracked change")));

const noComments = await parseDocx(fixture("no-comments.docx"));
assert.equal(noComments.summary.commentCount, 0);

const embedded = await parseDocx(fixture("embedded-object.docx"));
assert.ok(embedded.warnings.some((warning) => warning.includes("embedded object files")));
assert.ok(embedded.warnings.some((warning) => warning.includes("Embedded object relationship")));

const large = await parseDocx(fixture("large.docx"));
assert.equal(large.summary.blockCount, 1100);
assert.equal(large.summary.renderedBlockCount, 500);
assert.ok(large.warnings.some((warning) => warning.includes("rendered document list is capped")));

await assert.rejects(() => parseDocx(fixture("malformed.docx")), /valid .docx zip package/i);
await assert.rejects(() => parseDocx(fixture("encrypted.docx")), /Encrypted, password-protected, or legacy Word/i);

fs.rmSync(new URL("../.tmp-parser-test.mjs", import.meta.url));
console.log("DOCX parser fixture tests passed.");
