# GitHub Copilot Office Add-in — Hướng Dẫn Sử Dụng

**GitHub Copilot Office Add-in** là một tiện ích mở rộng (Add-in) cho Microsoft Office, tích hợp trợ lý AI GitHub Copilot trực tiếp vào **Word**, **Excel** và **PowerPoint**. Người dùng có thể trò chuyện với Copilot để đọc, chỉnh sửa và tạo nội dung tài liệu mà không cần rời khỏi ứng dụng Office.

---

## Mục lục

1. [Điều kiện tiên quyết (Prerequisites)](#1-điều-kiện-tiên-quyết-prerequisites)
2. [Các tính năng nổi bật](#2-các-tính-năng-nổi-bật)
3. [Kiến trúc & Cách hoạt động](#3-kiến-trúc--cách-hoạt-động)
4. [Hướng dẫn cài đặt](#4-hướng-dẫn-cài-đặt)
5. [Hướng dẫn sử dụng](#5-hướng-dẫn-sử-dụng)
6. [Danh mục công cụ (Tools Catalog)](#6-danh-mục-công-cụ-tools-catalog)
7. [Xử lý sự cố (Troubleshooting)](#7-xử-lý-sự-cố-troubleshooting)
8. [Câu hỏi thường gặp (FAQ)](#8-câu-hỏi-thường-gặp-faq)
9. [Đóng góp & Phát triển](#9-đóng-góp--phát-triển)
10. [Roadmap](#10-roadmap)

---

## 1. Điều kiện tiên quyết (Prerequisites)

### Phần mềm bắt buộc

| Phần mềm | Phiên bản tối thiểu | Link tải |
|-----------|---------------------|----------|
| **Node.js** | 20+ | [nodejs.org](https://nodejs.org/) |
| **Git** | Bất kỳ | [git-scm.com](https://git-scm.com/downloads) |
| **Microsoft Office** | Microsoft 365 hoặc Office 2019+ | — |
| **npm** | Đi kèm Node.js | — |

### Tài khoản & Quyền truy cập

- **Tài khoản GitHub** với **GitHub Copilot license** đang hoạt động (Individual, Business hoặc Enterprise).
- Đã đăng nhập GitHub Copilot CLI (`github-copilot` package) trên máy.

### Hệ điều hành hỗ trợ

| Hệ điều hành | Trạng thái |
|---------------|------------|
| macOS (Intel & Apple Silicon) | ✅ Hỗ trợ đầy đủ |
| Windows 10/11 | ✅ Hỗ trợ đầy đủ |
| Linux | ❌ Chưa hỗ trợ (Office Desktop không khả dụng) |

### Ứng dụng Office hỗ trợ

- **Microsoft Word** — Đọc, chỉnh sửa, định dạng tài liệu
- **Microsoft Excel** — Đọc/ghi dữ liệu, tạo biểu đồ, định dạng ô
- **Microsoft PowerPoint** — Tạo slide, chỉnh sửa nội dung, quản lý speaker notes

---

## 2. Các tính năng nổi bật

### 🤖 Trò chuyện AI trực tiếp trong Office
- Giao diện chat tích hợp ngay trong task pane của Word / Excel / PowerPoint.
- Hỗ trợ gửi cả **văn bản** và **hình ảnh** (paste hoặc đính kèm).

### 📝 Thao tác tài liệu thông minh
- **Word**: Đọc cấu trúc, tìm kiếm & thay thế, chèn bảng, định dạng văn bản, chèn nội dung tại vị trí con trỏ.
- **Excel**: Đọc/ghi dữ liệu ô, tạo biểu đồ (cột, đường, tròn, scatter...), định dạng ô, tạo Named Range.
- **PowerPoint**: Tạo slide lập trình từ code (PptxGenJS), chụp ảnh slide, quản lý speaker notes, nhân bản slide.

### 🔄 Đa mô hình AI
- Chọn mô hình AI từ danh sách khả dụng (Claude Sonnet 4.5, 4.6, v.v.) ngay trên giao diện.
- Tự động phát hiện và liệt kê các mô hình từ Copilot CLI.

### 🌐 Web Fetch
- Lấy nội dung từ URL bên ngoài và chuyển sang markdown, hỗ trợ nghiên cứu trực tiếp từ tài liệu Office.

### 💬 Quản lý phiên hội thoại
- Lưu lịch sử hội thoại theo từng ứng dụng Office (Word, Excel, PowerPoint).
- Khôi phục phiên cũ, tạo phiên mới bất cứ lúc nào.

### 🖥️ Tray App
- Ứng dụng chạy nền trên system tray (Windows) hoặc menu bar (macOS).
- Khởi động/dừng server, mở trang web quản trị trực tiếp từ tray icon.

### 🎨 Dark Mode
- Tự động nhận diện chế độ sáng/tối của hệ thống và trình duyệt.

### 🔒 Quản lý quyền (Permission)
- Hệ thống permission cho phép kiểm soát quyền đọc/ghi/shell khi Copilot thực thi các tool.

---

## 3. Kiến trúc & Cách hoạt động

### Sơ đồ kiến trúc tổng quan

```
┌─────────────────────────────────────────────────────────┐
│                   Microsoft Office                       │
│  ┌──────────┐  ┌──────────┐  ┌─────────────────────┐   │
│  │   Word   │  │  Excel   │  │     PowerPoint      │   │
│  └────┬─────┘  └────┬─────┘  └──────────┬──────────┘   │
│       │              │                   │               │
│       └──────────────┼───────────────────┘               │
│                      │                                   │
│              ┌───────▼────────┐                          │
│              │  Office Add-in │                          │
│              │  (Task Pane)   │                          │
│              │  React + Fluent│                          │
│              └───────┬────────┘                          │
└──────────────────────┼──────────────────────────────────┘
                       │ HTTPS (port 52390)
                       │
         ┌─────────────▼─────────────┐
         │     Express HTTPS Server  │
         │   (server.js / Vite HMR)  │
         │                           │
         │  ┌─────────────────────┐  │
         │  │  REST API Endpoints │  │
         │  │  /api/hello         │  │
         │  │  /api/upload-image  │  │
         │  │  /api/fetch         │  │
         │  │  /api/log           │  │
         │  └─────────────────────┘  │
         │                           │
         │  ┌─────────────────────┐  │
         │  │  WebSocket Proxy    │  │
         │  │  /api/copilot (WSS) │  │
         │  └──────────┬──────────┘  │
         └─────────────┼─────────────┘
                       │ stdio (JSON-RPC)
                       │
            ┌──────────▼──────────┐
            │  GitHub Copilot CLI │
            │  (@github/copilot)  │
            │                     │
            │  - Session mgmt    │
            │  - Model selection  │
            │  - Tool execution   │
            └──────────┬──────────┘
                       │
                       │ HTTPS API
                       │
            ┌──────────▼──────────┐
            │  GitHub Copilot     │
            │  Cloud Service      │
            │  (AI Models)        │
            └─────────────────────┘
```

### Luồng hoạt động (Flow)

```
Người dùng nhập tin nhắn vào Chat
        │
        ▼
React UI (ChatInput) ─── gửi prompt + attachments
        │
        ▼
WebSocket Client ─── kết nối WSS tới /api/copilot
        │
        ▼
WebSocket Proxy (copilotProxy.js)
        │ spawn child process
        ▼
Copilot CLI (--server --stdio)
        │ JSON-RPC qua stdin/stdout
        ▼
Copilot Cloud ─── xử lý AI, trả về kết quả
        │
        ▼
Stream events về WebSocket Client
        │
        ▼
React UI hiển thị kết quả (streaming)
        │
        ▼
Nếu Copilot gọi Tool (ví dụ: get_document_content)
        │
        ▼
Tool Handler thực thi qua Office.js API
        │
        ▼
Kết quả trả về cho Copilot → tiếp tục xử lý
```

### Thành phần chính

| Thành phần | File/Thư mục | Mô tả |
|------------|-------------|--------|
| **Express Server** | `src/server.js` | Server HTTPS phục vụ frontend + API (dev mode có Vite HMR) |
| **Production Server** | `src/server-prod.js` | Server phục vụ file tĩnh cho bản build production |
| **WebSocket Proxy** | `src/copilotProxy.js` | Proxy WebSocket ↔ Copilot CLI (stdio), xử lý LSP message framing |
| **React Frontend** | `src/ui/` | Giao diện chat, sử dụng Fluent UI React v9 |
| **Office Tools** | `src/ui/tools/` | Các tool tương tác với tài liệu Office qua Office.js API |
| **Tray App** | `src/tray/main.js` | Ứng dụng Electron chạy nền trên system tray |
| **Manifest** | `manifest.xml` | Khai báo add-in với Office (hosts, permissions, UI entry points) |

---

## 4. Hướng dẫn cài đặt

### 4.1. Clone mã nguồn & cài đặt dependencies

```bash
git clone <repository-url>
cd ghc-office-addin-fork
npm install
```

### 4.2. Đăng ký Add-in với Office

Script này sẽ tạo chứng chỉ SSL tự ký (self-signed), trust certificate và đăng ký manifest với Office.

**macOS:**
```bash
./register.sh
```

**Windows (PowerShell chạy quyền Administrator):**
```powershell
.\register.ps1
```

> ⚠️ Trên macOS, hệ thống có thể yêu cầu nhập mật khẩu để trust certificate vào Keychain.

### 4.3. Khởi chạy ứng dụng

#### Cách 1: Tray App (khuyến nghị)

```bash
npm run start:tray
```

Sau khi chạy, biểu tượng GitHub Copilot sẽ xuất hiện trên **system tray** (Windows) hoặc **menu bar** (macOS).

#### Cách 2: Development Server (có Hot Reload)

```bash
npm run dev
```

Server phát triển sẽ chạy trên `https://localhost:52390` với hot reload qua Vite.

#### Cách 3: Production Server

```bash
npm run build
npm run start
```

### 4.4. Build Installer (tùy chọn)

Tạo bộ cài đặt cho người dùng cuối:

```bash
# Build cho platform hiện tại
npm run build:installer

# Build riêng cho macOS (.dmg)
npm run build:installer:mac

# Build riêng cho Windows (.exe)
npm run build:installer:win
```

### 4.5. Thêm Add-in vào Office

1. Đảm bảo server/tray app đang chạy (kiểm tra biểu tượng trên tray/menu bar).
2. Mở **Word**, **Excel** hoặc **PowerPoint**.
   > ⚠️ Nếu ứng dụng đã mở trước khi đăng ký, hãy **đóng hoàn toàn và mở lại**.
3. Vào **Insert** → **Add-ins** → **My Add-ins**.
4. Tìm và chọn **GitHub Copilot**.
5. Task pane sẽ mở ra với giao diện chat — bạn đã sẵn sàng sử dụng!

---

## 5. Hướng dẫn sử dụng

### 5.1. Giao diện chính

Khi mở Add-in, bạn sẽ thấy giao diện gồm:

| Khu vực | Mô tả |
|---------|--------|
| **Header Bar** | Chọn mô hình AI, tạo phiên mới, xem lịch sử, bật/tắt debug mode |
| **Message List** | Hiển thị hội thoại giữa bạn và Copilot (hỗ trợ markdown, code blocks) |
| **Chat Input** | Nhập câu hỏi hoặc yêu cầu, đính kèm hình ảnh |

### 5.2. Cách sử dụng cơ bản

#### Trò chuyện với Copilot

1. Nhập câu hỏi hoặc yêu cầu vào ô chat phía dưới.
2. Nhấn **Enter** hoặc nút **Gửi** (biểu tượng mũi tên).
3. Copilot sẽ phản hồi theo thời gian thực (streaming).

#### Đính kèm hình ảnh

- **Paste** hình ảnh từ clipboard trực tiếp vào khung chat.
- Copilot có thể phân tích hình ảnh và sử dụng thông tin từ đó.

#### Chọn mô hình AI

- Click vào dropdown mô hình trên Header Bar.
- Chọn mô hình phù hợp (ví dụ: Claude Sonnet 4.5, Claude Sonnet 4.6).

#### Quản lý phiên hội thoại

- **Tạo phiên mới**: Click nút "New" trên Header Bar.
- **Xem lịch sử**: Click nút "History" để xem các phiên trước đó.
- **Khôi phục phiên**: Chọn một phiên từ danh sách lịch sử để xem lại.

### 5.3. Ví dụ sử dụng theo ứng dụng

#### Microsoft Word

| Yêu cầu mẫu | Mô tả |
|-------------|--------|
| *"Tóm tắt nội dung tài liệu này"* | Copilot đọc toàn bộ tài liệu và tạo bản tóm tắt |
| *"Thêm một bảng 3 cột ở vị trí con trỏ"* | Chèn bảng HTML tại selection |
| *"Tìm và thay thế 'ABC' bằng 'XYZ'"* | Tìm kiếm & thay thế nội dung |
| *"Bôi đậm đoạn text đang chọn"* | Áp dụng định dạng bold cho selection |
| *"Viết lại phần 'Introduction' cho ngắn gọn hơn"* | Đọc section, viết lại và cập nhật |

#### Microsoft Excel

| Yêu cầu mẫu | Mô tả |
|-------------|--------|
| *"Cho tôi xem tổng quan workbook"* | Liệt kê sheets, ranges, charts |
| *"Tạo biểu đồ cột từ dữ liệu A1:C10"* | Tạo chart từ data range |
| *"Định dạng hàng header thành bold, nền xanh"* | Áp dụng formatting cho cells |
| *"Tính tổng cột B và ghi vào B11"* | Ghi công thức/giá trị vào ô |
| *"Tìm tất cả ô chứa 'Error' và thay bằng 'OK'"* | Tìm kiếm & thay thế trong cells |

#### Microsoft PowerPoint

| Yêu cầu mẫu | Mô tả |
|-------------|--------|
| *"Tạo một bài thuyết trình 5 slide về AI"* | Tạo nhiều slide với nội dung |
| *"Cho xem nội dung slide 3"* | Đọc text content của slide |
| *"Chụp ảnh slide 1 cho tôi xem"* | Capture slide dưới dạng PNG |
| *"Thêm speaker notes cho slide 2"* | Ghi/cập nhật ghi chú thuyết trình |
| *"Nhân bản slide 4 ra cuối"* | Copy slide sang vị trí mới |

### 5.4. Mẹo sử dụng hiệu quả

1. **Luôn bắt đầu bằng overview** — Yêu cầu Copilot xem tổng quan tài liệu trước khi chỉnh sửa để nó hiểu context.
2. **Chỉnh sửa có chọn lọc** — Dùng các tool phẫu thuật (find & replace, insert at selection) thay vì thay toàn bộ nội dung.
3. **Gửi hình ảnh** — Paste ảnh thiết kế mẫu để Copilot tạo slide/bảng theo mẫu.
4. **PowerPoint: Dùng code** — Tool `add_slide_from_code` rất mạnh, cho phép tạo slide với layout phức tạp qua PptxGenJS API.
5. **Excel: Format sau data** — Ghi dữ liệu trước → định dạng → tạo biểu đồ → đặt named range.

---

## 6. Danh mục công cụ (Tools Catalog)

### Word Tools

| Tool | Mô tả |
|------|--------|
| `get_document_overview` | Xem tổng quan cấu trúc: số từ, heading, bảng, danh sách |
| `get_document_content` | Đọc toàn bộ nội dung HTML của tài liệu |
| `get_document_section` | Đọc nội dung một section theo tên heading |
| `set_document_content` | Thay thế toàn bộ nội dung body bằng HTML mới |
| `get_selection` | Lấy nội dung đang chọn dạng OOXML |
| `get_selection_text` | Lấy text thuần của vùng đang chọn |
| `insert_content_at_selection` | Chèn HTML tại vị trí con trỏ (before/after/replace) |
| `find_and_replace` | Tìm kiếm & thay thế (hỗ trợ case sensitivity, whole word) |
| `insert_table` | Chèn bảng có header styling và grid/striped |
| `apply_style_to_selection` | Áp dụng định dạng: bold, italic, underline, font size, màu sắc |

### PowerPoint Tools

| Tool | Mô tả |
|------|--------|
| `get_presentation_overview` | Xem tổng quan: số slide, preview nội dung |
| `get_presentation_content` | Đọc text từ slides (hỗ trợ phân trang) |
| `get_slide_image` | Chụp slide thành ảnh PNG |
| `get_slide_notes` | Đọc speaker notes |
| `set_presentation_content` | Thêm text box hoặc tạo slide mới |
| `add_slide_from_code` | Tạo slide bằng code PptxGenJS (mạnh nhất) |
| `clear_slide` | Xóa tất cả shapes trên slide |
| `update_slide_shape` | Cập nhật text của shape đã có |
| `set_slide_notes` | Thêm/sửa speaker notes |
| `duplicate_slide` | Nhân bản slide |

### Excel Tools

| Tool | Mô tả |
|------|--------|
| `get_workbook_overview` | Tổng quan: sheets, ranges, named ranges, charts |
| `get_workbook_info` | Liệt kê tên sheets và sheet đang active |
| `get_workbook_content` | Đọc giá trị & công thức từ ô/range |
| `set_workbook_content` | Ghi mảng 2D giá trị vào ô |
| `get_selected_range` | Đọc cells đang chọn |
| `set_selected_range` | Ghi giá trị vào vùng đang chọn |
| `find_and_replace_cells` | Tìm & thay thế trong cells |
| `insert_chart` | Tạo biểu đồ (column, bar, line, pie, area, scatter, doughnut) |
| `apply_cell_formatting` | Định dạng ô: bold, màu, border, number format, alignment |
| `create_named_range` | Tạo named range cho công thức |

### Tool chung

| Tool | Mô tả |
|------|--------|
| `web_fetch` | Lấy nội dung URL và chuyển sang markdown |

---

## 7. Xử lý sự cố (Troubleshooting)

### Add-in không xuất hiện trong Office

| Bước | Hành động |
|------|-----------|
| 1 | Kiểm tra server đang chạy: truy cập `https://localhost:52390` |
| 2 | Kiểm tra biểu tượng Copilot trên system tray / menu bar |
| 3 | Đóng **hoàn toàn** ứng dụng Office và mở lại |
| 4 | Chạy lại script đăng ký (`register.sh` hoặc `register.ps1`) |
| 5 | Xóa Office cache và thử lại |

### Lỗi chứng chỉ SSL

```bash
# macOS: Chạy lại register script
./register.sh

# Windows: Chạy PowerShell với quyền Admin
.\register.ps1
```

Nếu vẫn lỗi, trust thủ công file `certs/localhost.pem` vào Keychain (macOS) hoặc Certificate Store (Windows).

### Service không khởi động sau khi cài installer

- **Windows**: Kiểm tra Task Scheduler cho task "CopilotOfficeAddin".
- **macOS**: Kiểm tra launchctl:
  ```bash
  launchctl list | grep copilot
  cat /tmp/copilot-office-addin.log
  ```

### WebSocket không kết nối

1. Đảm bảo port `52390` không bị chiếm bởi ứng dụng khác.
2. Kiểm tra firewall không chặn kết nối localhost.
3. Mở Developer Tools (F12) trong Office task pane để xem log lỗi.

### Copilot không phản hồi

1. Kiểm tra đã đăng nhập GitHub Copilot CLI:
   ```bash
   npx @github/copilot --version
   ```
2. Đảm bảo license Copilot còn hiệu lực.
3. Kiểm tra kết nối internet.

---

## 8. Câu hỏi thường gặp (FAQ)

**Q: Add-in có hoạt động với Office Online (web) không?**
> Hiện tại add-in được thiết kế cho Office Desktop (Word, Excel, PowerPoint trên Windows và macOS). Office Online chưa được hỗ trợ chính thức.

**Q: Dữ liệu tài liệu có được gửi lên server bên ngoài không?**
> Nội dung tài liệu được xử lý cục bộ qua Office.js API. Khi bạn yêu cầu Copilot phân tích, nội dung sẽ được gửi tới GitHub Copilot Cloud Service (tương tự khi dùng Copilot trong VS Code).

**Q: Tôi có thể dùng mô hình AI nào?**
> Các mô hình khả dụng phụ thuộc vào license GitHub Copilot của bạn. Add-in sẽ tự động liệt kê danh sách mô hình từ Copilot CLI.

**Q: Port 52390 bị trùng thì sao?**
> Port `52390` là cố định (strict port) do manifest.xml khai báo. Bạn cần giải phóng port này trước khi chạy server. Kiểm tra bằng: `lsof -i :52390` (macOS) hoặc `netstat -ano | findstr 52390` (Windows).

**Q: Làm sao gỡ cài đặt (unregister)?**
> ```bash
> # macOS
> ./unregister.sh
> 
> # Windows (PowerShell)
> .\unregister.ps1
> ```

---

## 9. Đóng góp & Phát triển

### Cấu trúc dự án

```
├── src/
│   ├── server.js            # Dev server (Vite + Express)
│   ├── server-prod.js       # Production server (static files)
│   ├── copilotProxy.js      # WebSocket proxy ↔ Copilot CLI
│   └── ui/                  # React frontend
│       ├── App.tsx           # Component chính
│       ├── components/       # ChatInput, MessageList, HeaderBar, ...
│       ├── lib/              # WebSocket client, permission, logging
│       └── tools/            # Office.js tool handlers
├── dist/                    # Built frontend assets
├── certs/                   # SSL certificates (localhost)
├── assets/                  # Icons, tray images
├── installer/               # Installer scripts (macOS, Windows)
├── manifest.xml             # Office add-in manifest
├── register.sh / .ps1       # Setup scripts
└── unregister.sh / .ps1     # Cleanup scripts
```

### Các lệnh hữu ích cho developer

| Lệnh | Mô tả |
|-------|--------|
| `npm run dev` | Khởi chạy dev server với hot reload |
| `npm run build` | Build frontend cho production |
| `npm run start` | Chạy production server |
| `npm run start:tray` | Chạy Electron tray app |
| `npm run build:installer` | Build installer cho platform hiện tại |
| `npm run build:installer:mac` | Build macOS `.dmg` installer |
| `npm run build:installer:win` | Build Windows `.exe` installer |

### Công nghệ sử dụng

| Công nghệ | Vai trò |
|-----------|---------|
| **React 18** | UI framework |
| **Fluent UI React v9** | Design system (Microsoft) |
| **Vite** | Build tool & dev server |
| **Express** | HTTPS server |
| **Electron** | Tray app desktop |
| **WebSocket** | Giao tiếp real-time với Copilot CLI |
| **Office.js** | API tương tác với tài liệu Office |
| **@github/copilot-sdk** | SDK kết nối GitHub Copilot |
| **vscode-jsonrpc** | JSON-RPC protocol cho LSP messaging |
| **PptxGenJS** | Tạo slide PowerPoint bằng code |

---

## 10. Roadmap

### ✅ Đã hoàn thành
- [x] Tích hợp cơ bản với Word, Excel, PowerPoint
- [x] Giao diện chat với streaming response
- [x] Hỗ trợ đính kèm hình ảnh
- [x] Tray app (Electron) cho macOS & Windows
- [x] Tool system đầy đủ cho cả 3 ứng dụng Office
- [x] Lịch sử phiên hội thoại
- [x] Dark mode
- [x] Chọn mô hình AI
- [x] Tạo slide từ code (PptxGenJS)
- [x] Biểu đồ Excel (7 loại chart)
- [x] Web fetch tích hợp

### 🔄 Đang phát triển
- [ ] Code signing cho installer (macOS & Windows)
- [ ] Standalone installer chính thức (không cần Node.js)
- [ ] Cải thiện hệ thống permission chi tiết hơn

### 📋 Kế hoạch tương lai
- [ ] Hỗ trợ OneNote (manifest đã khai báo host Notebook)
- [ ] Office Online (Web) support
- [ ] Đa ngôn ngữ (i18n) cho giao diện
- [ ] Plugin/Extension system cho custom tools
- [ ] Tích hợp GitHub Issues & Pull Requests từ tài liệu Office
- [ ] Template library — bộ sưu tập mẫu slide/tài liệu có sẵn
- [ ] Collaboration features — chia sẻ phiên chat giữa nhiều người dùng
- [ ] Auto-update mechanism cho tray app
- [ ] Telemetry & analytics dashboard (opt-in)
- [ ] Hỗ trợ file đính kèm ngoài hình ảnh (PDF, CSV, v.v.)

---

> **Ghi chú**: Tài liệu này được viết dựa trên phiên bản `1.0.0` của dự án. Nội dung có thể thay đổi theo các bản cập nhật mới.
