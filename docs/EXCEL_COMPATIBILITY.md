# Ma trận tương thích Excel

Tài liệu này là bản kiểm kê phạm vi đã có trong mã nguồn và test của gói
Spreadsheet tại thời điểm hiện tại. “Có” nghĩa là engine có API/đường chạy và
được test; “Một phần” nghĩa là chỉ hỗ trợ một tập con hoặc cần tích hợp thêm;
“Chưa có” nghĩa là chưa có implementation spreadsheet tương ứng. Đây không
phải chứng nhận tương thích Excel hay Google Sheets, và không tuyên bố Excel
parity.

Các liên kết trong bảng trỏ trực tiếp tới source hoặc test làm căn cứ cho từng
nhận định.

## Ma trận phạm vi

| Phạm vi | Trạng thái | Những gì đang có | Căn cứ trong mã/test | Khác biệt hoặc giới hạn quan trọng |
| --- | --- | --- | --- | --- |
| Data model và view | Có | Mô hình một worksheet với cột có `id`, tiêu đề, kiểu text/number/boolean, số dòng, ô input thưa; đọc raw input, computed value và display value. Có global/per-column search, typed filter, ordered stable multi-sort, versioned column-view state và view index. | [`model.ts`](../packages/spreadsheet/src/model.ts), [`view.ts`](../packages/spreadsheet/src/view.ts), [`column-view-state.ts`](../packages/spreadsheet/src/column-view-state.ts), [`view-index.ts`](../packages/spreadsheet/src/view-index.ts), [`model.test.ts`](../packages/spreadsheet/test/model.test.ts), [`column-view-state.test.ts`](../packages/spreadsheet/test/column-view-state.test.ts), [`view.test.ts`](../packages/spreadsheet/test/view.test.ts) | Chỉ là single-sheet model; chưa có workbook nhiều sheet, table object, pivot hay named range. View index phải được consumer invalidate/refresh khi dữ liệu đổi. |
| Formula syntax và references | Một phần | Công thức đồng bộ, không dùng `eval`/`Function`; toán tử số học, so sánh, nối chuỗi, unary, ngoặc; A1, absolute/mixed references, range, structured reference và một số whole-column dependency. Có literal lỗi `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#N/A`, `#CYCLE!`. | [`formula.ts`](../packages/spreadsheet/src/formula.ts), [`formula-authoring.ts`](../packages/spreadsheet/src/formula-authoring.ts), [`formula.test.ts`](../packages/spreadsheet/test/formula.test.ts), [`formula-authoring.test.ts`](../packages/spreadsheet/test/formula-authoring.test.ts) | Chưa có cross-sheet reference, named range, table semantics đầy đủ, dynamic array/spill, external link, 3D reference hay toàn bộ cú pháp Excel. Structured ref là cú pháp package (`[Column]`, `[@Column]`), không phải một table engine. |
| Hàm tích hợp và hàm mở rộng | Một phần | Registry tên không phân biệt hoa thường, register/unregister/has/list; custom function nhận scalar hoặc range có metadata shape và chỉ được đồng bộ. Built-in hiện có: `ABS`, `AND`, `AVERAGE`, `CONCAT`, `COUNTA`, `COUNT`, `DATE`, `DAY`, `IF`, `INDEX`, `LEFT`, `LEN`, `LOWER`, `MATCH`, `MAX`, `MID`, `MIN`, `MONTH`, `NOT`, `OR`, `RIGHT`, `ROUND`, `SUM`, `TODAY`, `TRIM`, `UPPER`, `XLOOKUP`, `YEAR`. | [`function-registry.ts`](../packages/spreadsheet/src/function-registry.ts), [`formula.ts`](../packages/spreadsheet/src/formula.ts), [`formula.test.ts`](../packages/spreadsheet/test/formula.test.ts), [`model.test.ts`](../packages/spreadsheet/test/model.test.ts) | `MATCH` là exact-only (match type khác 0/khuyết cho kết quả `#N/A`); `INDEX`/`XLOOKUP` là MVP exact/1D theo các kiểm tra shape hiện có. Registry không tự override built-in; override phải explicit. Không có toàn bộ thư viện hàm Excel/Sheets, macro hay async UDF. |
| Calculation, dependency và errors | Có | Dependency được thu từ AST, kể cả argument của custom function; model invalidates theo ô/cột phụ thuộc, lazy-compute và phát hiện cycle. Exception, kết quả không hợp lệ hoặc Promise từ custom function chuyển thành lỗi `#VALUE!` deterministic. Có chính sách aggregate cho lỗi và coercion được cấu hình. | [`formula.ts`](../packages/spreadsheet/src/formula.ts), [`model.ts`](../packages/spreadsheet/src/model.ts), [`aggregate.ts`](../packages/spreadsheet/src/aggregate.ts), [`formula.test.ts`](../packages/spreadsheet/test/formula.test.ts), [`model.test.ts`](../packages/spreadsheet/test/model.test.ts), [`aggregate.test.ts`](../packages/spreadsheet/test/aggregate.test.ts) | Không nhằm tái tạo toàn bộ coercion, precision, volatile calculation, iterative calculation hoặc error propagation của Excel. Registry mutation không tự quan sát; consumer phải gọi `recalculateAll()`/`invalidateAll()` để thay đổi có hiệu lực xác định. |
| Editing, history, fill và structural | Có | Cell edit có transaction, coalesce duplicate/no-op, bounded undo/redo và selection metadata. `WorkbookHistory`/`SpreadsheetSession` xếp chung cell command với snapshot/structural command. Smart fill hỗ trợ công thức, dãy số và ngày ISO. Find/Replace hỗ trợ raw/computed search, raw replacement plan có giới hạn và demo UI. Structural engine insert/delete row/column bất biến, dịch công thức A1, xử lý `#REF!`, và trả `beforePayload`/`afterPayload`. | [`history.ts`](../packages/spreadsheet/src/history.ts), [`workbook-history.ts`](../packages/spreadsheet/src/workbook-history.ts), [`spreadsheet-session.ts`](../packages/spreadsheet/src/spreadsheet-session.ts), [`find-replace.ts`](../packages/spreadsheet/src/find-replace.ts), [`fill.ts`](../packages/spreadsheet/src/fill.ts), [`structural.ts`](../packages/spreadsheet/src/structural.ts), [`spreadsheet-session.test.ts`](../packages/spreadsheet/test/spreadsheet-session.test.ts), [`find-replace.test.ts`](../packages/spreadsheet/test/find-replace.test.ts), [`structural.test.ts`](../packages/spreadsheet/test/structural.test.ts) | Snapshot không được deep-clone; structural operations materialize raw rows once per operation. Storybook structural buttons chưa dùng session mới. Replace All gồm cả source rows đang bị view filter ẩn và bị chặn nếu result `truncated`. Cell compensation là best-effort; snapshot atomicity thuộc callback consumer. |
| Formatting và validation | Một phần | Format số, phần trăm, tiền tệ, ngày, text theo `Intl`; validation required, range số, độ dài, one-of, regex, ISO date range và custom predicate; formula policy raw/skip/computed. Conditional-format engine có rule theo số, text, empty, duplicate, formula error và custom style patch. | [`format.ts`](../packages/spreadsheet/src/format.ts), [`validation.ts`](../packages/spreadsheet/src/validation.ts), [`conditional-format.ts`](../packages/spreadsheet/src/conditional-format.ts), [`format.test.ts`](../packages/spreadsheet/test/format.test.ts), [`validation.test.ts`](../packages/spreadsheet/test/validation.test.ts), [`conditional-format.test.ts`](../packages/spreadsheet/test/conditional-format.test.ts) | Chưa có mô hình formatting giàu như font/fill/border/alignment theo workbook, data-validation dropdown UI, protection hay toàn bộ format code Excel. Conditional formatting là engine renderer-independent; việc vẽ thuộc consumer/UI. |
| Import/export và persistence | Một phần | CSV parser/stringifier có quote, delimiter, type inference, giới hạn kích thước, policy formula text/reject và export raw/computed; JSON snapshot v1 sparse có validate, limit, migrate và metadata primitive; workbook I/O restore được model mới. | [`csv.ts`](../packages/spreadsheet/src/csv.ts), [`workbook-io.ts`](../packages/spreadsheet/src/workbook-io.ts), [`persistence.ts`](../packages/spreadsheet/src/persistence.ts), [`csv.test.ts`](../packages/spreadsheet/test/csv.test.ts), [`workbook-io.test.ts`](../packages/spreadsheet/test/workbook-io.test.ts), [`persistence.test.ts`](../packages/spreadsheet/test/persistence.test.ts) | Không có XLSX/ODS, rich workbook package, styles/formulas cross-sheet, chart/pivot, binary import hay collaboration persistence. CSV chọn raw formula hoặc computed value; đây không phải round-trip Excel workbook. Snapshot restore hiện tạo model mới và lazy rebuild cache. |
| UI, accessibility và performance | Một phần | Storybook harness có grid editing, formula bar/completion/diagnostics, global controls và popover khi click tên cột để search/filter/sort; Shift-click giữ multi-sort. Có undo/redo, fill, conditional-format demo, CSV/JSON controls; các control chính có label, dialog focus, live status và alert. Có `SpreadsheetViewIndex` cho cache search, model dependency invalidation và benchmark lặp lại cho 100k dòng/1M cell. | [`spreadsheet.stories.tsx`](../packages/spreadsheet/src/spreadsheet.stories.tsx), [`column-view-state.ts`](../packages/spreadsheet/src/column-view-state.ts), [`view-index.ts`](../packages/spreadsheet/src/view-index.ts), [`performance.mjs`](../packages/spreadsheet/bench/performance.mjs), [`integration.test.ts`](../packages/spreadsheet/test/integration.test.ts), [`view-index.test.ts`](../packages/spreadsheet/test/view-index.test.ts) | Đây là harness tích hợp chứ chưa phải spreadsheet UI/component API công khai. Header canvas chưa có đường mở popover bằng keyboard độc lập; icon/menu và dialog controls dùng pointer/focus hiện có. Chưa có bộ a11y chuyên biệt, worker calculation, ngân sách hiệu năng chuẩn hóa theo hardware, collaboration hay mobile input contract. View-index không tự biết mọi mutation và không hỗ trợ đổi số cột sau khi khởi tạo. |

## Hàm hiện có và registry

Danh sách trên được lấy từ built-in dispatch và collision set trong
[`formula.ts`](../packages/spreadsheet/src/formula.ts) và
[`function-registry.ts`](../packages/spreadsheet/src/function-registry.ts),
không phải danh sách hàm Excel được suy đoán. Tên built-in/custom được chuẩn
hóa không phân biệt hoa thường. Đăng ký trùng tên bị từ chối; built-in chỉ bị
thay thế khi consumer yêu cầu `overrideBuiltIn`. Hàm custom được truyền
`FormulaFunctionArgument`, có thể là scalar hoặc range với `rows`/`columns`, và
không nhận AST nội bộ. Hàm phải trả về `FormulaValue` đồng bộ; Promise/thenable,
exception hoặc giá trị ngoài tập này cho `#VALUE!`.

## Những khác biệt cần biết khi so với Excel/Google Sheets

- Mục tiêu hiện tại là engine một sheet, synchronous và có thể nhúng; không phải
  clone Excel/Sheets. Không nên dùng ma trận này để tuyên bố parity.
- A1 và structured references được hỗ trợ trong phạm vi package, nhưng chưa có
  cross-sheet, named ranges, table lifecycle, dynamic arrays/spills hoặc
  external links.
- Lookup là MVP exact: `MATCH` chưa có approximate/sorted modes; `XLOOKUP` chỉ
  exact và range một chiều; không có toàn bộ tùy chọn/coercion của Excel.
- Ngày được biểu diễn qua `FormulaValue` (thường là số hoặc chuỗi ISO trong
  input); chưa có date-system toggle, timezone/workbook locale semantics như
  một file Excel đầy đủ. Format hiển thị dùng `Intl` locale do consumer chọn.
- Formula parser dùng cú pháp và separator giới hạn của engine, không bao phủ
  toàn bộ function names, array constants, lambdas, macros, add-ins hay locale
  separators của Excel/Sheets.
- CSV/JSON là các boundary rõ ràng của package. Export công thức raw và export
  computed value là hai chế độ khác nhau; không giữ toàn bộ workbook feature
  như XLSX.
- Structural edits và history là các engine callback/payload; ứng dụng sở hữu
  model mutation, persistence và atomicity của side effects.

## Đề xuất tier để tránh phình phạm vi

### Core MVP

Giữ ổn định các phần đã có: single-sheet model, A1/mixed/structured refs,
dependency/cycle/error deterministic, built-in hiện tại, synchronous registry,
view/filter/sort, CSV và JSON snapshot, validation/format cơ bản, fill,
structural payload transform và các history adapter. Ưu tiên tài liệu API,
invariant và test hơn là thêm hàng trăm hàm tương thích một phần.

### Phase 2

Chỉ mở rộng khi có use case cụ thể: cross-sheet/named range hoặc table model,
locale-aware formula parsing, richer date/error/coercion semantics, thêm lookup
mode có specification rõ, worker adapter cho calculation/view vượt ngưỡng, và adapter import/export workbook riêng (ví dụ XLSX)
để không kéo dependency vào core. Dynamic arrays nên có thiết kế spill/dependency
riêng trước khi đưa vào model hiện tại.

### Optional

Để ngoài core những phần dễ tạo chi phí bảo trì lớn: chart/pivot, rich Excel
formatting, macros/add-ins, volatile/async UDF, comments/protection,
collaboration, distributed calculation và UI spreadsheet hoàn chỉnh.
Các tính năng này phù hợp adapter hoặc package riêng với contract hiệu năng và
security độc lập.

## Phạm vi kiểm chứng

Ma trận được đối chiếu với các source và test được link trực tiếp ở trên. Khi
implementation thay đổi, cập nhật tài liệu cùng test tương ứng; nếu một claim
không còn được chứng minh bởi source/test, hãy hạ trạng thái hoặc xóa claim
thay vì diễn giải thành tương thích Excel rộng hơn.
