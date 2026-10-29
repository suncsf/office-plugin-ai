import Database from "better-sqlite3";
const db = new Database("data/app.db", { readonly: true });
const rows = db
  .prepare("SELECT id, role, content, thinking, tool_calls FROM messages ORDER BY id DESC LIMIT 12")
  .all();
for (const r of rows) {
  console.log(
    JSON.stringify({
      id: r.id,
      role: r.role,
      contentHead: (r.content || "").slice(0, 30),
      thinkingLen: (r.thinking || "").length,
      thinkingHead: (r.thinking || "").slice(0, 60),
      hasToolCalls: !!r.tool_calls,
    })
  );
}
