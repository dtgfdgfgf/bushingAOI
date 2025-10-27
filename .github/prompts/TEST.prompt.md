---
description: New prompt created from chat session
mode: agent
tools: ['extensions', 'codebase', 'usages', 'vscodeAPI', 'problems', 'changes', 'testFailure', 'terminalSelection', 'terminalLastCommand', 'openSimpleBrowser', 'fetch', 'findTestFiles', 'searchResults', 'githubRepo', 'runCommands', 'runTasks', 'editFiles', 'runNotebooks', 'search', 'new']
model: GPT-5 (Preview)
---
<user>
請幫我看一下我的md 並給予評價
</user>
<assistant>

</assistant>
<user>
@GitHubCopilot Enable: "Enable GPT-5 (Preview) for all clients"
</user>
<assistant>
以下是精簡評價與建議：

優點
- 結構清楚，分區合理（Coding/Docs/DB）。
- SUPPORT_EXPERT 條目具體，重視可維護性、效能與測試。
- MySQL 指南點到關鍵（InnoDB、索引、連線池）。

可加強
- 仍有未替換變數：{{bushingAOI}}、{{project-description}}、{{query_patterns}}。
- SUPPORT_LEVEL 僅定義 EXPERT，層級不完整與適用情境未說明。
- 規範缺少強制力詞彙（MUST/SHOULD/MAY），不利一致性與審核。
- MySQL 部分不足：隔離級別、遷移流程、觀測性、備份還原、權限與安全。
- 文檔流程缺少驗收門檻（如需連結 PR、更新版本/ADR/變更紀錄）。

優先建議
1) 明確化範疇與強制力（RFC 2119 詞彙）。  
2) 補齊 SUPPORT_STANDARD/MINIMAL 並說明何時適用。  
3) 擴充 MySQL 實務：索引策略細則、遷移、連線池預設、慢查與備份。  
4) 文件更新流程加入必填項（README/CHANGELOG/ADR/破壞性變更標示）。  
5) 新增「模板變數說明」區，避免遺漏占位符。

建議修正版（可直接覆蓋原檔）
````markdown
# AI Rules for {{bushingAOI}}

{{project-description}}

Version: 0.2
Last Updated: 2025-08-31
Scope: Applies to AI assistants and contributors in this repository.

Conventions
- Use RFC 2119 keywords: MUST / SHOULD / MAY for enforceability.
- Prefer clarity and maintainability; justify deviations.

Template Variables
- {{bushingAOI}}: Project name.
- {{project-description}}: One-paragraph summary of the system.
- {{query_patterns}}: Typical query shapes used to guide indexing.

## CODING_PRACTICES

### Guidelines for SUPPORT_LEVEL

#### SUPPORT_EXPERT
- Favor elegant, maintainable solutions over verbose code.
- Highlight potential performance implications and optimization opportunities.
- Frame solutions within broader architectural context; suggest alternatives when appropriate.
- Comment on the "why", not the "what".
- Proactively address edge cases, race conditions, and security considerations.
- When debugging, provide targeted diagnostics, not shotgun approaches.
- Suggest comprehensive testing strategies (mocking, organization, coverage).

#### SUPPORT_STANDARD
- Provide clear, production-ready solutions with minimal dependencies.
- Include basic error handling, logging, and input validation.
- Offer at least one alternative approach and note trade-offs.

#### SUPPORT_MINIMAL
- Provide concise code snippets or pointers; defer architecture decisions.
- Call out assumptions and risks explicitly.
- Link to relevant docs or prior art.

## DOCUMENTATION

### DOC_UPDATES
- MUST update relevant documentation in /docs when modifying features.
- MUST keep README.md in sync with new capabilities and usage changes.
- MUST maintain changelog entries in CHANGELOG.md (tag breaking changes).
- SHOULD update or create ADRs for significant decisions and link from PRs.

## DATABASE

### Guidelines for SQL

#### MYSQL
- MUST use InnoDB storage engine for transactions and foreign key constraints.
- Isolation: default REPEATABLE READ; adjust per-transaction if needed for hot paths.
- Indexing (based on {{query_patterns}}):
  - Prefer composite and covering indexes; order by selectivity and predicate usage.
  - Ensure foreign-key columns are indexed; avoid redundant duplicates.
  - Avoid leading wildcards; consider prefix indexes for large varchars.
  - Review EXPLAIN before merge for non-trivial queries.
- Queries:
  - Use prepared statements and parameterized queries.
  - Paginate large result sets; prefer keyset pagination for deep pages.
  - Limit SELECT columns to what is needed.
- Migrations:
  - Use migration tooling (e.g., Flyway/Liquibase), forward-only, reversible when feasible.
  - Zero/low-downtime patterns (backfill, dual-write, flip, cleanup).
- Connection pooling:
  - MUST use pooling. Suggested defaults: maxPoolSize 10–20, connectionTimeout 30s, idleTimeout 1–5m, maxLifetime 30–60m.
- Observability:
  - Enable slow query log (e.g., >500ms) and track top offenders.
  - Collect query plans for regressions.
- Reliability & Safety:
  - Backups: daily incremental + weekly full; quarterly restore drills.
  - Least-privilege DB user; rotate credentials; no SUPER on app accounts.
  - Enforce NOT NULL, CHECKs (where supported), and explicit defaults.

````

需要我直接套用上述修正版到檔案，或再依你的實際技術棧與 {{query_patterns}} 具體化內容嗎？
</assistant>
<user>
Analyze this codebase to generate or update `.github/copilot-instructions.md` for guiding AI coding agents.

Focus on discovering the essential knowledge that would help an AI agents be immediately productive in this codebase. Consider aspects like:
- The "big picture" architecture that requires reading multiple files to understand - major components, service boundaries, data flows, and the "why" behind structural decisions
- Critical developer workflows (builds, tests, debugging) especially commands that aren't obvious from file inspection alone
- Project-specific conventions and patterns that differ from common practices
- Integration points, external dependencies, and cross-component communication patterns

Source existing AI conventions from `**/{.github/copilot-instructions.md,AGENT.md,AGENTS.md,CLAUDE.md,.cursorrules,.windsurfrules,.clinerules,.cursor/rules/**,.windsurf/rules/**,.clinerules/**,README.md}` (do one glob search).

Guidelines (read more at https://aka.ms/vscode-instructions-docs):
- If `.github/copilot-instructions.md` exists, merge intelligently - preserve valuable content while updating outdated sections
- Write concise, actionable instructions (~20-50 lines) using markdown structure
- Include specific examples from the codebase when describing patterns
- Avoid generic advice ("write tests", "handle errors") - focus on THIS project's specific approaches
- Document only discoverable patterns, not aspirational practices
- Reference key files/directories that exemplify important patterns

Update `.github/copilot-instructions.md` for the user, then ask for feedback on any unclear or incomplete sections to iterate.
</user>
<assistant>

</assistant>