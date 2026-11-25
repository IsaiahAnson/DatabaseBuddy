// Unit tests for the step logic. They run without Pipedream: OneDrive, OpenAI, Notion
// and the data store are replaced with small fakes. Run with: node --test test/*.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { loadStep, bind } from "./load-step.mjs";

const step01 = await loadStep("01_comprehensive_onedrive_processor.js");
const step02 = await loadStep("02_clean_and_deduplicate_data.js");
const step03 = await loadStep("03_create_specialized_it_databases.js");
const step04 = await loadStep("04_populate_specialized_databases.js");
const step05 = await loadStep("05_normalize_database_records.js");
const step06 = await loadStep("06_enhance_individual_records.js");
const step07 = await loadStep("07_update_notion_pages_with_enhanced_content.js");

const fakeStep = () => ({ exports: {}, export(name, value) { this.exports[name] = value; } });

function fakeDataStore() {
  const map = new Map();
  return {
    map,
    async get(key) { return map.get(key); },
    async set(key, value) { map.set(key, value); },
  };
}

const ROUTER_CATEGORIES = Object.keys(bind(step04).getDatabaseMapping());

// ---------- categories ----------

test("step 01 prompt names exactly the ten categories step 04 routes", async () => {
  let request;
  const ctx = bind(step01, {
    openai: {
      async createChatCompletion(args) {
        request = args;
        return { generated_message: { content: JSON.stringify({ knowledgeCategory: "security" }) } };
      },
    },
  });
  const result = await ctx.enhanceWithAI("text", "doc.txt", {}, fakeStep());
  const prompt = request.data.messages[1].content;
  const listed = prompt.match(/exactly one of these names: (.+?)\. Pick/)[1].split(", ");
  assert.deepEqual(listed, ROUTER_CATEGORIES);
  assert.equal(ROUTER_CATEGORIES.length, 10);
  assert.deepEqual(request.data.response_format, { type: "json_object" });
  assert.equal(request.data.model, "gpt-4o");
  assert.ok(request.$, "the step's $ is passed to the OpenAI app");
  assert.equal(result.success, true);
  assert.equal(result.analysis.knowledgeCategory, "Security");
});

test("step 01 reports an OpenAI error as a failed analysis", async () => {
  const ctx = bind(step01, {
    openai: { async createChatCompletion() { throw new Error("401 Unauthorized"); } },
  });
  const result = await ctx.enhanceWithAI("text", "doc.txt", {}, fakeStep());
  assert.equal(result.success, false);
  assert.match(result.analysis.extractionNotes, /401/);
});

test("step 04 normalizeCategory maps canonical names, aliases and longer labels", () => {
  const ctx = bind(step04);
  for (const name of ROUTER_CATEGORIES) {
    assert.equal(ctx.normalizeCategory(name.toUpperCase()), name);
  }
  assert.equal(ctx.normalizeCategory("networking"), "Network Infrastructure");
  assert.equal(ctx.normalizeCategory("Microsoft Intune"), "MDM");
  assert.equal(ctx.normalizeCategory("Windows Autopilot"), "MDM");
  assert.equal(ctx.normalizeCategory("Azure AD Connect"), "User Management");
  assert.equal(ctx.normalizeCategory("Azure"), "Cloud Services");
  assert.equal(ctx.normalizeCategory("Azure Virtual Desktop"), "VDI");
  assert.equal(ctx.normalizeCategory("BitLocker"), "Security");
  assert.equal(ctx.normalizeCategory("Identity and Access Management"), "User Management");
  // Whole words only: "application" must not match inside "applicationsupport".
  assert.equal(ctx.normalizeCategory("applicationsupport"), "applicationsupport");
  assert.equal(ctx.normalizeCategory("Development"), "Development");
  assert.equal(ctx.normalizeCategory(""), "");
});

test("step 01 matchCategory returns canonical names and keeps other labels", () => {
  const ctx = bind(step01);
  assert.equal(ctx.matchCategory("backup recovery"), "Backup Recovery");
  assert.equal(ctx.matchCategory(" Microsoft Intune "), "Microsoft Intune");
  assert.equal(ctx.matchCategory(undefined), "Uncategorized");
});

// ---------- step 01: dedupe key, record mapping, scan ----------

test("step 01 skips an item only when the stored version matches", async () => {
  const store = fakeDataStore();
  const ctx = bind(step01, { data_store: store });
  const item = { id: "A1", eTag: "\"{X},1\"", lastModifiedDateTime: "2026-01-01T00:00:00Z" };
  assert.equal(ctx.dedupeKey(item), "onedrive:A1");
  assert.equal(await ctx.isUnchanged(item), false);
  store.map.set("onedrive:A1", { version: item.eTag });
  assert.equal(await ctx.isUnchanged(item), true);
  assert.equal(await ctx.isUnchanged({ ...item, eTag: "\"{X},2\"" }), false);
  assert.equal(ctx.isSameVersion({ version: "2026-01-01T00:00:00Z" }, { id: "B", lastModifiedDateTime: "2026-01-01T00:00:00Z" }), true);
  assert.equal(ctx.isSameVersion({ version: null }, { id: "C" }), false);
});

test("step 01 toRecord produces the shared record shape", () => {
  const ctx = bind(step01);
  const record = ctx.toRecord({
    metadata: { id: "A1", eTag: "e1", name: "Intune enrollment.md", webUrl: "https://example.com/doc" },
    extractedContent: { success: true, content: "Steps to enroll a device." },
    aiAnalysis: {
      success: true,
      analysis: {
        documentType: "Technical Guide",
        keyTopics: ["Intune", "Enrollment"],
        importantInfo: "How to enroll a device.",
        knowledgeCategory: "mdm",
      },
    },
  });
  assert.deepEqual(record, {
    id: "A1",
    eTag: "e1",
    title: "Intune enrollment.md",
    category: "MDM",
    content_type: "Technical Guide",
    tags: ["Intune", "Enrollment"],
    summary: "How to enroll a device.",
    content: "Steps to enroll a device.",
    source: "OneDrive",
    webUrl: "https://example.com/doc",
  });
});

test("step 01 scanFolder follows paging, matches excludes exactly and skips unchanged files", async () => {
  const pages = {
    "/items/root/children": {
      value: [
        { id: "f-temp", name: "temp", folder: {} },
        { id: "f-templates", name: "Templates", folder: {} },
        { id: "a", name: "a.txt", eTag: "1", file: { mimeType: "text/plain" } },
      ],
      "@odata.nextLink": "https://graph.example/next",
    },
    "https://graph.example/next": {
      value: [{ id: "b", name: "b.txt", eTag: "1", file: { mimeType: "text/plain" } }],
    },
    "/items/f-templates/children": {
      value: [{ id: "t", name: "t.txt", eTag: "1", file: { mimeType: "text/plain" } }],
    },
  };
  const requested = [];
  const store = fakeDataStore();
  store.map.set("onedrive:a", { version: "1" });
  const ctx = bind(step01, {
    scanSubfolders: true,
    data_store: store,
    microsoft_onedrive: {
      async httpRequest({ url }) {
        const key = url.split("?")[0];
        requested.push(key);
        return pages[key];
      },
    },
  });
  const stats = { skippedUnchanged: 0, folderErrors: [] };
  const files = await ctx.scanFolder("root", [], ["temp"], ["text/plain"], 10, stats, true);
  assert.deepEqual(files.map(f => f.id), ["t", "b"]);
  assert.equal(stats.skippedUnchanged, 1);
  assert.ok(!requested.includes("/items/f-temp/children"), "the excluded folder is not scanned");
  assert.ok(requested.includes("https://graph.example/next"), "the next page is requested");
});

test("step 01 scanFolder fails on the top folder and records subfolder errors", async () => {
  const failing = {
    microsoft_onedrive: { async httpRequest() { throw new Error("401 InvalidAuthenticationToken"); } },
    data_store: fakeDataStore(),
    scanSubfolders: true,
  };
  const stats = { skippedUnchanged: 0, folderErrors: [] };
  await assert.rejects(
    bind(step01, failing).scanFolder("root", [], [], ["text/plain"], 5, stats, true),
    /401/
  );
  await bind(step01, failing).scanFolder("sub", [], [], ["text/plain"], 5, stats);
  assert.equal(stats.folderErrors.length, 1);
});

// ---------- step 02 ----------

test("step 02 batch ids are stable and stored records are not returned again", async () => {
  const store = fakeDataStore();
  const props = {
    data_store: store,
    include_onedrive: true,
    include_excel: false,
    include_onenote: false,
    similarity_threshold: 80,
    batch_size: 10,
    max_records: 3,
    use_persistent_storage: true,
  };
  const a = { id: "a", content: "intune enrollment guide for windows laptops" };
  const b = { id: "b", content: "bitlocker recovery key escrow procedure" };
  const c = { id: "c", content: "cisco switch vlan trunk configuration" };

  const ctx = bind(step02, props);
  assert.equal(ctx.batchId([a, b]), ctx.batchId([a, b]));
  assert.notEqual(ctx.batchId([a, b]), ctx.batchId([a, c]));

  const first = await step02.run.call({ ...ctx, onedrive_data: [a, b] }, { $: fakeStep() });
  assert.deepEqual(first.final_records.map(r => r.id), ["a", "b"]);

  const again = await step02.run.call({ ...ctx, onedrive_data: [a, b] }, { $: fakeStep() });
  assert.deepEqual(again.final_records, []);

  const next = await step02.run.call({ ...ctx, onedrive_data: [c] }, { $: fakeStep() });
  assert.deepEqual(next.final_records.map(r => r.id), ["c"]);
});

// ---------- step 03 ----------

test("step 03 sends a type object for every property under initial_data_source", async () => {
  let sent;
  const ctx = bind(step03, {
    title: "Security",
    properties: [
      JSON.stringify({ name: "Name", type: "title" }),
      JSON.stringify({ name: "Source URL", type: "url" }),
      JSON.stringify({ name: "Category", type: "select", select: { options: [{ name: "Security" }] } }),
    ],
    notion: { async createDatabase(body) { sent = body; return { id: "db" }; } },
  });
  await step03.run.call(ctx, { $: fakeStep() });
  assert.equal(sent.properties, undefined);
  assert.deepEqual(sent.initial_data_source.properties, {
    "Name": { type: "title", title: {} },
    "Source URL": { type: "url", url: {} },
    "Category": { type: "select", select: { options: [{ name: "Security" }] } },
  });

  const bad = bind(step03, { title: "x", properties: [JSON.stringify({ name: "No type" })], notion: {} });
  await assert.rejects(step03.run.call(bad, { $: fakeStep() }), /must have a 'name' and 'type'/);
});

// ---------- step 04 ----------

function fakeNotion({ existing = {} } = {}) {
  const calls = { created: [], updated: [], queried: [] };
  return {
    calls,
    async queryDataSource(dataSourceId, params) {
      calls.queried.push({ dataSourceId, params });
      const sourceId = params.filter.rich_text.equals;
      return { results: existing[sourceId] ? [{ id: existing[sourceId] }] : [] };
    },
    async createPage(body) {
      calls.created.push(body);
      return { id: `page-${calls.created.length}` };
    },
    async updatePage(pageId, body) {
      calls.updated.push({ pageId, body });
      return { id: pageId };
    },
    async retrievePage(pageId) {
      return { id: pageId, properties: { Name: { type: "title" } } };
    },
  };
}

const sampleRecord = {
  id: "A1",
  eTag: "e1",
  title: "Intune enrollment.md",
  category: "Microsoft Intune",
  content_type: "Technical Guide",
  tags: ["Intune", "Enrollment, Windows"],
  summary: "How to enroll a device.",
  content: "Steps to enroll a device.",
  source: "OneDrive",
  webUrl: "https://example.com/doc",
};

test("step 04 writes the record fields to the documented properties", () => {
  const props = bind(step04).createPageProperties({ ...sampleRecord, category: "MDM" });
  assert.equal(props["Name"].title[0].text.content, "Intune enrollment.md");
  assert.equal(props["Source ID"].rich_text[0].text.content, "A1");
  assert.equal(props["Source URL"].url, "https://example.com/doc");
  assert.equal(props["Category"].select.name, "MDM");
  assert.equal(props["Content Type"].select.name, "Technical Guide");
  assert.deepEqual(props["Tags"].multi_select, [{ name: "Intune" }, { name: "Enrollment  Windows" }]);
  assert.equal(props["Content Summary"].rich_text[0].text.content, "How to enroll a device.");
  assert.equal(bind(step04).createPageProperties({ title: "x" })["Content Type"].select, null);
});

test("step 04 creates a new page, updates an existing one and passes page ids on", async () => {
  const notion = fakeNotion({ existing: { B2: "page-existing" } });
  const ctx = bind(step04, {
    notion,
    mdmDb: "ds-mdm",
    securityDb: "ds-sec",
    cleanedData: [
      JSON.stringify(sampleRecord),
      JSON.stringify({ ...sampleRecord, id: "B2", title: "BitLocker.md", category: "Security" }),
    ],
  });
  const out = await step04.run.call(ctx, { $: fakeStep() });
  assert.deepEqual(notion.calls.queried[0], {
    dataSourceId: "ds-mdm",
    params: { filter: { property: "Source ID", rich_text: { equals: "A1" } }, page_size: 1 },
  });
  assert.equal(notion.calls.created.length, 1);
  assert.equal(notion.calls.created[0].parent.data_source_id, "ds-mdm");
  assert.equal(notion.calls.updated.length, 1);
  assert.equal(notion.calls.updated[0].pageId, "page-existing");
  assert.equal(notion.calls.updated[0].body.properties["Created Date"], undefined);
  assert.equal(out.created_count, 1);
  assert.equal(out.updated_count, 1);
  assert.equal(out.success, true);
  assert.deepEqual(out.records.map(r => [r.id, r.category, r.pageId]), [
    ["A1", "MDM", "page-1"],
    ["B2", "Security", "page-existing"],
  ]);
});

test("step 04 fails the run when no record reaches Notion", async () => {
  const ctx = bind(step04, {
    notion: fakeNotion(),
    cleanedData: [JSON.stringify({ ...sampleRecord, category: "Uncategorized" })],
  });
  await assert.rejects(step04.run.call(ctx, { $: fakeStep() }), /No pages were written/);
});

// ---------- step 05 ----------

test("step 05 keeps the record fields later steps read", async () => {
  const defaults = Object.fromEntries(
    Object.entries(step05.props).filter(([, p]) => p.default).map(([k, p]) => [k, p.default])
  );
  const ctx = bind(step05, {
    ...defaults,
    records: [JSON.stringify({ ...sampleRecord, category: "MDM", pageId: "page-1" })],
  });
  const out = await step05.run.call(ctx, { $: fakeStep() });
  const record = out.normalized_records[0];
  assert.equal(record.id, "A1");
  assert.equal(record.title, "Intune enrollment.md");
  assert.equal(record.category, "MDM");
  assert.equal(record.pageId, "page-1");
  assert.equal(record.summary, "How to enroll a device.");
  assert.equal(record.content_type, "Technical Guide");
  assert.equal(record.webUrl, "https://example.com/doc");
  assert.equal(record.metadata.pageId, undefined);
});

// ---------- steps 06 and 07 ----------

const aiReply = {
  enhanced_title: "Enrolling Windows Devices in Intune",
  detailed_summary: "A step-by-step guide to enrolling Windows devices.",
  categories: ["Endpoint"],
  suggested_tags: ["intune", "windows"],
  priority_level: "High",
  content_type: "Document",
};

test("step 06 writes the AI output into title, summary and tags", () => {
  const record = { ...sampleRecord, category: "MDM", pageId: "page-1" };
  const out = bind(step06, { model: "gpt-4o-mini" }).applyEnhancement(record, aiReply);
  assert.equal(out.title, "Enrolling Windows Devices in Intune");
  assert.equal(out.summary, "A step-by-step guide to enrolling Windows devices.");
  assert.deepEqual(out.tags, ["intune", "windows"]);
  assert.equal(out.category, "MDM");
  assert.equal(out.content_type, "Technical Guide");
  assert.equal(out.pageId, "page-1");
  assert.equal(out.original_title, "Intune enrollment.md");
  assert.equal(out.priority_level, "High");
  assert.equal(out.enhanced_title, undefined);
});

test("step 06 asks for JSON and fails the run when every record fails", async () => {
  let request;
  const ok = bind(step06, {
    model: "gpt-4o-mini",
    records: [JSON.stringify(sampleRecord)],
    openai: {
      async createChatCompletion(args) {
        request = args;
        return { generated_message: { content: JSON.stringify(aiReply) } };
      },
    },
  });
  const out = await step06.run.call(ok, { $: fakeStep() });
  assert.deepEqual(request.data.response_format, { type: "json_object" });
  assert.equal(out.enhanced_records[0].title, aiReply.enhanced_title);

  const failing = bind(step06, {
    model: "gpt-4o-mini",
    records: [JSON.stringify(sampleRecord)],
    openai: { async createChatCompletion() { throw new Error("429 rate limit"); } },
  });
  await assert.rejects(step06.run.call(failing, { $: fakeStep() }), /Enhancement failed for all 1 records/);

  const truncated = bind(step06, {
    model: "gpt-4o-mini",
    records: [JSON.stringify(sampleRecord)],
    openai: { async createChatCompletion() { return { generated_message: { content: "{\"enhanced_title\": \"Cut" } }; } },
  });
  await assert.rejects(step06.run.call(truncated, { $: fakeStep() }), /not valid JSON/);
});

test("step 06 keeps the record's id and pageId whatever the model returns", () => {
  const out = bind(step06, { model: "gpt-4o-mini" })
    .applyEnhancement({ ...sampleRecord, pageId: "page-1" }, { ...aiReply, id: "x", pageId: "y" });
  assert.equal(out.id, "A1");
  assert.equal(out.pageId, "page-1");
});

test("contract: a step 06 record changes the title, summary and tags in step 07", async () => {
  const enhanced = bind(step06, { model: "gpt-4o-mini" })
    .applyEnhancement({ ...sampleRecord, category: "MDM", pageId: "page-1" }, aiReply);
  const ctx = bind(step07, { updateTitles: true, addSummaries: true });
  const existingPage = { properties: { Name: { type: "title" } } };
  const props = await ctx.buildPageProperties(enhanced, existingPage);
  assert.equal(props["Name"].title[0].text.content, "Enrolling Windows Devices in Intune");
  assert.equal(props["Content Summary"].rich_text[0].text.content, aiReply.detailed_summary);
  assert.deepEqual(props["Tags"].multi_select, [{ name: "intune" }, { name: "windows" }]);
  assert.equal(props["Summary"], undefined);
});

test("step 07 never sends a parent, skips failed records and fails when nothing updates", async () => {
  assert.equal(step07.props.parentPageId, undefined);
  const enhanced = bind(step06, { model: "gpt-4o-mini" })
    .applyEnhancement({ ...sampleRecord, pageId: "page-1" }, aiReply);
  const notion = fakeNotion();
  notion.appendBlock = async () => ({});
  const ctx = bind(step07, {
    notion,
    updateTitles: true,
    addSummaries: true,
    replaceContent: false,
    enhancedRecords: [
      JSON.stringify(enhanced),
      JSON.stringify({ ...sampleRecord, pageId: "page-2", enhancement_error: "429" }),
    ],
    createdPages: [JSON.stringify({ id: "page-1" }), JSON.stringify({ id: "page-2" })],
  });
  const out = await step07.run.call(ctx, { $: fakeStep() });
  assert.equal(notion.calls.updated.length, 1);
  assert.equal(notion.calls.updated[0].body.parent, undefined);
  assert.equal(out.updatedPages.length, 1);
  assert.equal(out.skipped.length, 1);

  const broken = fakeNotion();
  broken.retrievePage = async () => { throw new Error("object_not_found"); };
  const failing = bind(step07, {
    notion: broken,
    updateTitles: true,
    addSummaries: true,
    enhancedRecords: [JSON.stringify(enhanced)],
  });
  await assert.rejects(step07.run.call(failing, { $: fakeStep() }), /No pages were updated/);
});
