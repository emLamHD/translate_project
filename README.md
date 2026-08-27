# Deterministic NO-AI SOP Translator

`TRANS-BE-001-C1` cung cấp nền tảng translation memory và formatter DOCX chạy
hoàn toàn xác định trước. Runtime không gọi dịch vụ bên ngoài, không tải model,
không OCR và không tự đoán bản dịch còn thiếu.

## NO-AI có nghĩa là gì?

Với `--execution-profile no-ai`:

- chỉ dùng exact match hoặc normalized template đã được con người phê duyệt;
- chỉ chấp nhận provenance `human_approved` hoặc `owner_manual` với `approved: true`;
- chặn outbound socket trong toàn bộ pipeline;
- không khởi tạo hay tải AI/ML model;
- không gọi Claude, OpenAI, Google Translate, OPUS/Marian/Transformers hoặc dịch vụ ngoài;
- không dùng fuzzy match làm bản dịch chính thức;
- câu chưa có bản dịch hợp lệ nhận trạng thái `MANUAL_TRANSLATION_REQUIRED`;
- không chèn marker thiếu vào DOCX trừ khi bật `--show-missing-markers`;
- coverage dưới 100% luôn tạo file `_no-ai_INCOMPLETE.docx`.

`RUNTIME_AI_USED: NO` chỉ mô tả runtime của package này. Dự án được phát triển
với Codex hỗ trợ. Hai reference riêng tư hiện tại từng được Claude tạo và chỉ là
`claude_silver_reference`, không phải bản dịch do con người phê duyệt. Draft
`blind_01_bilingual_draft.docx` từng dùng Google Translate và được phân loại
`AI_DERIVED_LEGACY_DRAFT`; nó không được dùng làm input cho NO-AI pipeline.

Không tuyên bố output đạt độ chính xác 100% hoặc sẵn sàng phát hành khi coverage
chưa đủ và chưa có Owner/chuyên gia duyệt nội dung.

## Import Translation Memory đã được con người phê duyệt

CSV tối thiểu:

```csv
unit_id,source_text,target_text,approved_by
unit-001,Cân chính xác 5 g mẫu.,Accurately weigh 5 g of the sample.,REVIEWER
```

```bash
PYTHONPATH=src python -m translator.cli tm-import \
  --input private_tm/approved_translations.csv \
  --source vi --target en \
  --approved-by "OWNER_OR_REVIEWER" \
  --output private_tm/vi_en_approved.json \
  --audit-report private_reports/tm_import.audit.json
```

Importer kiểm tra hash nguồn, duplicate/conflict và token bất biến. Nó không sửa
target. Cache Google hoặc Claude reference không bao giờ tự được nâng cấp.

## Chạy tài liệu

```bash
PYTHONPATH=src python -m translator.cli translate \
  private_samples/blind_01_source.doc \
  --execution-profile no-ai \
  --source vi --target en \
  --format-profile clean \
  --translation-memory private_tm/vi_en_approved.json \
  --output-dir outputs/no-ai
```

Nếu chưa có TM đã phê duyệt, bỏ `--translation-memory`. Pipeline vẫn kiểm tra
extraction, skip policy, formatter và cấu trúc nhưng kết quả mang nhãn `INCOMPLETE`.

## Formatting profiles

- `preserve` (mặc định): giữ layout/style nguồn; chỉ style target được chèn.
- `clean`: chuẩn hóa nhẹ spacing target, keep-with-next theo style heading có
  sẵn và margin ô bảng; không thay nội dung hoặc page setup.
- `etech-sop`: đọc thông số xác định trước từ JSON Owner phê duyệt qua
  `--format-config`. Không tự thiết kế hoặc suy luận bằng AI.

Heading chỉ được nhận diện từ style OOXML hiện có. Formatter không viết lại,
dịch, sửa chính tả, đổi số liệu, xóa hàng/cột hay thay đổi nội dung.

## Visual và embedded objects

`VISUAL_CONTENT_POLICY: PRESERVED_UNCHANGED_AND_SKIPPED`

Paragraph chứa `w:drawing`, `w:pict`, VML, OLE, equation, textbox hoặc object
nhúng bị bỏ qua toàn bộ. Pipeline không OCR, không dịch nhãn trong ảnh/chart,
không nhân bản relationship và không thay đổi resize/crop/anchor/wrap/alt text.
Structural QA so sánh counts, relationship hashes, media/chart/embedding hashes
và canonical object hashes trước/sau.

## Chứng minh zero network/model usage

Mỗi report riêng tư ghi `runtime_ai_used: false`, `external_translation_calls: 0`,
`outbound_document_content_calls: 0`, và `models_loaded: []`. `NetworkGuard` chặn
đường mở socket trong process. QA kiểm tra module AI/ML bị cấm không được load.
Tests quét source để bảo đảm không có Google endpoint, `urllib.request` hoặc OCR path.

## Quyền riêng tư

`.gitignore` loại private samples, DOC/DOCX/PDF, render image, TM, cache, private
report, `.env`, secrets và model artifacts. Không đặt dữ liệu khách hàng trong
fixtures, logs hoặc commit. Repository chỉ lưu test synthetic tự tạo ở runtime.

## Kiểm thử

```bash
PYTHONPATH=src pytest
ruff check src tests
mypy src
```

Synthetic integration bao phủ paragraph, heading, numbering, bảng/merged cell,
header/footer, ảnh, text+drawing, equation, mocked chart và embedded object.
