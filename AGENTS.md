# 妯北涔℃潙鎸叴椤圭洰 鈥?Codex 鏅鸿兘浣撻厤缃?
> Codex AI Agent 閰嶇疆 鈥?鏈枃浠跺畾涔変簡 Codex 鍦ㄨ椤圭洰涓殑琛屼负妯″紡鍜岃兘鍔涜竟鐣?
## 椤圭洰韬唤

- **椤圭洰鍚嶇О:** 涓婃捣鏉炬睙鍖烘í灞变埂鏉戞尟鍏存敼閫犻」鐩?- **椤圭洰绫诲瀷:** 椤跺眰绛栧垝 / 鍖轰綅鐮旂┒ / 瀵规爣鍒嗘瀽 / 鎷涘晢涓庤繍钀ユ柟妗?- **宸ヤ綔鐩綍:** `C:\Users\Lenovo\Documents\shanghai-hengshan-rural-revitalization`
- **榛樿妯″瀷:** deepseek-v4-flash

## 鍙敤鑳藉姏

### 鏍稿績鑳藉姏
| 鑳藉姏 | 瀹炵幇鏂瑰紡 | 鐘舵€?|
|---|---|---|
| 鏂囨。鐢熸垚 (Word) | python-docx | 鉁?|
| 琛ㄦ牸澶勭悊 (Excel) | openpyxl + pandas | 鉁?|
| 婕旂ず鏂囩 (PPT) | pptxgenjs | 鉁?|
| 缃戦〉娴忚/鎴浘 | Playwright Chromium | 鉁?|
| 妗岄潰杩滅▼鎺у埗 | Computer Use (named pipe) | 鉁?|
| AI 缂栫爜鏅鸿兘浣?| CodeWhale | 鉁?|
| 缃戠粶浠ｇ悊 | 127.0.0.1:51008 | 鉁?|

### 宸插畨瑁呭伐鍏风储寮?| 宸ュ叿 | 瀹夎浣嶇疆 | 鐢ㄩ€?|
|---|---|---|
| CodeWhale | npm global | DeepSeek 缂栫爜鏅鸿兘浣撶粓绔?|
| Playwright | bundled runtime | 娴忚鍣ㄨ嚜鍔ㄥ寲銆佹埅鍥?|
| python-docx | bundled Python | Word 鏂囨。璇诲啓 |
| openpyxl | bundled Python | Excel 琛ㄦ牸澶勭悊 |
| pptxgenjs | bundled Node.js | PPT 婕旂ず鏂囩鐢熸垚 |
| codex-proxy-switcher-win | ~/.codex/skills/ | 浠ｇ悊鍚姩鍣?(GUI) |

## 宸ヤ綔娴佽鍒?1. **瑙勫垝鍏堣:** 澶嶆潅浠诲姟鍏堣緭鍑虹粨鏋勫寲璁″垝锛岀‘璁ゅ悗鍐嶆墽琛?2. **璇佹嵁浼樺厛:** 宸ュ叿杈撳嚭浼樹簬鐚滄祴锛岄獙璇佹槸浠诲姟鐨勪竴閮ㄥ垎
3. **鐢ㄦ埛鎰忓浘浼樺厛:** 褰撳墠璇锋眰浼樺厛浜庡巻鍙茶蹇嗗拰鏃㈡湁閰嶇疆
4. **杈撳嚭鏍￠獙:** 鐢熸垚鏂囦欢鍚庤嚜鍔ㄩ獙璇佸唴瀹瑰畬鏁存€?5. **閿欒澶勭悊:** DeepSeek 闄愭祦/瓒呮椂鏃惰嚜鍔ㄩ檷绾ч噸璇?
## 瀹夊叏杈圭晫
- 鎵€鏈夎緭鍑烘枃浠朵繚瀛樺埌椤圭洰 output/ 鐩綍
- 涓嶄慨鏀圭郴缁熺骇閰嶇疆
- 涓嶈闂潪鎺堟潈鐩綍
- 浠ｇ悊閰嶇疆: 127.0.0.1:51008

---

*鐢?Codex 鑷姩鐢熸垚 鈥?2026-06-09*

---

Use installed `superpowers-*` skills whenever the task matches `obra/superpowers` workflows such as brainstorming, planning, execution, TDD, debugging, code review, git worktree isolation, branch finishing, verification, or skill authoring.

Codex has native skills and subagents. Prefer those primitives directly instead of translating the workflow into long ad-hoc prompts.

Default document output language is Simplified Chinese for plans, specs, reviews, summaries, postmortems, design docs, ADRs, and status updates.

When creating a new durable document-style file without an explicit user-provided path, prefer the repository's documentation directory. Use `docs/` by default, unless the repository already clearly uses another documentation directory such as `doc/`, `spec/`, or `specs/`.

When creating a new document-style file without an explicit user-provided name, prefer a concise Simplified Chinese filename that matches the document's actual purpose, such as `瀹炴柦璁″垝.md`, `浠ｇ爜璇勫.md`, `闂鎺掓煡.md`, `鎺ュ彛璁捐.md`, `鏁版嵁缁撴瀯璁捐.md`, `琛ㄧ粨鏋勮璁?md`, `Redis璁捐.md`, `S3璁捐.md`, or `瀛楁璇存槑.md`, unless the repository already uses an English naming convention.

If the document is specifically about Redis or S3, prefer literal names like `Redis璁捐.md` and `S3璁捐.md` over broader names such as cache design or object storage design.

Write document content in Simplified Chinese and keep it direct, concrete, and easy for Chinese-speaking teammates to read. Prefer plain language over heavy jargon. If a technical term is necessary, keep it accurate and add a brief explanation when that helps readability.

Keep source code, commands, file paths, URLs, logs, stack traces, environment variable names, schema identifiers, and existing English API terms in their original language unless the user explicitly asks for translation.

When the current workspace is already a linked worktree or detached HEAD, do not create another worktree blindly. Run setup and baseline checks first. If branch creation, push, or PR actions are blocked by the host sandbox, commit the work and hand off via the host application's native branch or local-handoff controls instead of claiming the workflow is impossible.

When the user is still exploring requirements, constraints, trade-offs, overall design, or explicitly asks to think first, prefer `superpowers-brainstorming`.

When the user already wants a concrete document deliverable such as an implementation plan, interface design, request/response contract, data structure, table structure, Redis design, S3 design, field descriptions, or an OpenAPI-style skeleton, prefer `superpowers-writing-plans` directly. Only use `superpowers-brainstorming` first if key decisions are still unresolved.

If the user provides a `TODO.md`, backlog file, worklist, or similar task list and asks for a plan, treat that file as the requirement source. Preserve the list order, start from the first actionable item, and only expand forward when the first item is too small to form a reasonable standalone work slice.

If design or planning documents are produced, explicitly capture the relevant data structures, interface contracts, field definitions, validation rules, naming conventions, retention rules, and error cases when they matter to the task.

Respect explicit user language overrides and repository-local documentation conventions.

