(() => {
  "use strict";

  const MM_TO_POINTS = 72 / 25.4;
  const BODY_FONT = "Times New Roman";
  const BODY_SIZE = 14;
  const KNOWN_TYPES = [
    "CƯƠNG LĨNH CHÍNH TRỊ", "ĐIỀU LỆ ĐẢNG", "CHIẾN LƯỢC", "NGHỊ QUYẾT",
    "QUYẾT ĐỊNH", "QUY ĐỊNH", "QUY CHẾ", "CHỈ THỊ", "KẾT LUẬN", "HƯỚNG DẪN",
    "THÔNG BÁO", "THÔNG CÁO", "BÁO CÁO", "TỜ TRÌNH", "KẾ HOẠCH", "QUY HOẠCH",
    "CHƯƠNG TRÌNH", "ĐỀ ÁN", "PHƯƠNG ÁN", "DỰ ÁN", "CÔNG VĂN", "BIÊN BẢN",
    "TUYÊN BỐ", "LỜI KÊU GỌI", "GIẤY GIỚI THIỆU", "GIẤY CHỨNG NHẬN", "GIẤY ĐI ĐƯỜNG",
    "GIẤY NGHỈ PHÉP", "GIẤY MỜI", "PHIẾU CHUYỂN", "PHIẾU GỬI", "THƯ CÔNG"
  ];

  const byId = (id) => document.getElementById(id);
  const buttons = ["normalize", "title", "summary", "review"].map(byId);

  function setStatus(message, kind = "") {
    const node = byId("status");
    node.textContent = message;
    node.dataset.kind = kind;
  }

  function normalizeSpace(value) {
    return (value || "").replace(/\s+/g, " ").trim();
  }

  function isTypeTitle(text) {
    const upper = text.toLocaleUpperCase("vi-VN");
    return KNOWN_TYPES.some((type) => upper === type || upper.startsWith(`${type} `) || upper.startsWith(`${type}:`));
  }

  function formatBodyParagraph(paragraph) {
    paragraph.font.set({ name: BODY_FONT, size: BODY_SIZE, color: "#000000" });
    paragraph.alignment = Word.Alignment.justified;
    paragraph.firstLineIndent = 10 * MM_TO_POINTS;
    paragraph.spaceAfter = 6;
    paragraph.lineSpacing = 20;
  }

  function formatCentered(paragraph, size, bold = true) {
    paragraph.font.set({ name: BODY_FONT, size, color: "#000000", bold });
    paragraph.alignment = Word.Alignment.centered;
    paragraph.firstLineIndent = 0;
    paragraph.spaceAfter = 0;
  }

  function formatRecognizedParagraphs(paragraphs) {
    let titleCount = 0;
    let headingCount = 0;
    const typeIndex = paragraphs.findIndex((p) => isTypeTitle(normalizeSpace(p.text)));
    let summaryIndex = -1;
    if (typeIndex >= 0) {
      for (let i = typeIndex + 1; i < paragraphs.length; i += 1) {
        const candidate = normalizeSpace(paragraphs[i].text);
        if (!candidate) continue;
        if (candidate.length < 180 && !/^(Phần|Chương|Mục|Điều)\b/i.test(candidate)) summaryIndex = i;
        break;
      }
    }
    paragraphs.forEach((paragraph, index) => {
      const text = normalizeSpace(paragraph.text);
      if (!text || paragraph.tableNestingLevel > 0) return;

      if (text.toLocaleUpperCase("vi-VN") === "ĐẢNG CỘNG SẢN VIỆT NAM") {
        formatCentered(paragraph, 15, true);
        titleCount += 1;
        return;
      }

      if (index < 10 && /^(BAN CHẤP HÀNH|ĐẢNG BỘ|ĐẢNG UỶ|ĐẢNG ỦY|VĂN PHÒNG|ĐẢNG CỘNG SẢN)/i.test(text) && text === text.toLocaleUpperCase("vi-VN")) {
        formatCentered(paragraph, 14, true);
        titleCount += 1;
        return;
      }

      if (isTypeTitle(text)) {
        formatCentered(paragraph, 15, true);
        titleCount += 1;
        return;
      }

      if (index === summaryIndex) {
        paragraph.font.set({ name: BODY_FONT, size: 14, color: "#000000", bold: true });
        paragraph.alignment = Word.Alignment.centered;
        paragraph.firstLineIndent = 0;
        paragraph.spaceAfter = 6;
        return;
      }

      if (/^(PHẦN|CHƯƠNG|MỤC)\s+/i.test(text)) {
        formatCentered(paragraph, 14, true);
        headingCount += 1;
        return;
      }

      if (/^ĐIỀU\s+\d+\s*\./i.test(text)) {
        paragraph.font.set({ name: BODY_FONT, size: BODY_SIZE, color: "#000000", bold: true });
        paragraph.alignment = Word.Alignment.left;
        paragraph.firstLineIndent = 0;
        paragraph.spaceAfter = 6;
        headingCount += 1;
        return;
      }

      if (/^(Hà Nội|TP\.|Thành phố|Tỉnh|Thị xã|Quận|Huyện|Phường|Xã|Đặc khu).*ngày\s+\d+/i.test(text)) {
        paragraph.font.set({ name: BODY_FONT, size: BODY_SIZE, color: "#000000", italic: true });
        paragraph.alignment = Word.Alignment.centered;
        paragraph.firstLineIndent = 0;
        paragraph.spaceAfter = 0;
      }

      if (/^Số\s*[:.]?/i.test(text)) {
        paragraph.alignment = Word.Alignment.left;
        paragraph.firstLineIndent = 0;
      }

      if (/^(Nơi nhận|Kính gửi|Kính trình)\s*:/i.test(text)) {
        paragraph.font.set({ name: BODY_FONT, size: BODY_SIZE, color: "#000000" });
        paragraph.firstLineIndent = 0;
        paragraph.alignment = Word.Alignment.left;
      }

      if (/^(T\/M\s|KT\.\s|PHÓ\s|BÍ\sTHƯ|CHỦ\sTỊCH|GIÁM\sĐỐC|CHÁNH\sVĂN\sPHÒNG)/i.test(text) || /^(Nơi nhận:|Lưu:)/i.test(text)) {
        paragraph.firstLineIndent = 0;
        paragraph.alignment = Word.Alignment.left;
      }
    });
    return { titleCount, headingCount };
  }

  async function normalizeDocument() {
    setStatus("Đang chuẩn hóa định dạng…");
    await Word.run(async (context) => {
      const bodyParagraphs = context.document.body.paragraphs;
      bodyParagraphs.load("items/text,items/tableNestingLevel");
      const supportsPageSetup = Office.context.requirements.isSetSupported("WordApiDesktop", "1.3");
      let sections = null;
      if (supportsPageSetup) {
        sections = context.document.sections;
        sections.load("items");
      }
      await context.sync();

      let pageUpdated = false;
      if (sections) {
        sections.items.forEach((section) => {
          section.pageSetup.set({
            paperSize: "A4",
            orientation: "Portrait",
            topMargin: 20 * MM_TO_POINTS,
            bottomMargin: 20 * MM_TO_POINTS,
            leftMargin: 30 * MM_TO_POINTS,
            rightMargin: 15 * MM_TO_POINTS,
            mirrorMargins: false
          });
        });
        pageUpdated = true;
      }

      const items = bodyParagraphs.items;
      const titleIndex = items.findIndex((p) => isTypeTitle(normalizeSpace(p.text)));
      let bodyStart = titleIndex >= 0 ? titleIndex + 1 : 0;
      if (titleIndex >= 0) {
        for (let i = titleIndex + 1; i < items.length; i += 1) {
          const candidate = normalizeSpace(items[i].text);
          if (!candidate) continue;
          if (candidate.length < 180 && !/^(Phần|Chương|Mục|Điều)\b/i.test(candidate)) bodyStart = i + 1;
          break;
        }
      }
      let bodyEnd = items.length;
      for (let i = bodyStart; i < items.length; i += 1) {
        const line = normalizeSpace(items[i].text);
        if (/^(Nơi nhận\s*:|T\/M\s|KT\.\s|Lưu\s*:)/i.test(line)) { bodyEnd = i; break; }
      }
      let formatted = 0;
      items.forEach((paragraph, index) => {
        if (paragraph.tableNestingLevel > 0 || !normalizeSpace(paragraph.text)) return;
        if (index >= bodyStart && index < bodyEnd) {
          formatBodyParagraph(paragraph);
          formatted += 1;
        }
      });
      const recognized = formatRecognizedParagraphs(bodyParagraphs.items);
      await context.sync();

      const pageMessage = pageUpdated ? "Đã đặt A4 và lề trang." : "Bản Word này chưa hỗ trợ tự đặt khổ giấy/lề; phần nội dung đã được định dạng.";
      setStatus(`${pageMessage} Đã chuẩn hóa ${formatted} đoạn nội dung (giữ nguyên phần đầu, nơi nhận, chữ ký và bảng); nhận diện ${recognized.titleCount} dòng thể thức và ${recognized.headingCount} đề mục.`, "success");
    });
  }

  async function formatSelection(kind) {
    setStatus("Đang định dạng vùng chọn…");
    await Word.run(async (context) => {
      const paragraphs = context.document.getSelection().paragraphs;
      paragraphs.load("items/text");
      await context.sync();
      if (!paragraphs.items.length) {
        setStatus("Hãy chọn một dòng văn bản trong Word trước.", "error");
        return;
      }

      paragraphs.items.forEach((paragraph) => {
        paragraph.font.set({
          name: BODY_FONT,
          size: kind === "title" ? 15 : 14,
          color: "#000000",
          bold: true
        });
        paragraph.alignment = Word.Alignment.centered;
        paragraph.firstLineIndent = 0;
        paragraph.spaceAfter = 0;
        paragraph.lineSpacing = 20;
      });
      await context.sync();
      setStatus(kind === "title" ? "Đã định dạng tên loại văn bản (15 pt, đậm, căn giữa)." : "Đã định dạng trích yếu (14 pt, đậm, căn giữa).", "success");
    });
  }

  async function reviewDocument() {
    setStatus("Đang rà soát các thành phần có thể nhận diện…");
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();
      const text = paragraphs.items.map((p) => p.text || "").join("\n");
      const lines = paragraphs.items.map((p) => normalizeSpace(p.text || ""));
      const checks = [
        ["Tiêu đề ĐẢNG CỘNG SẢN VIỆT NAM", /ĐẢNG\s+CỘNG\s+SẢN\s+VIỆT\s+NAM/i.test(text)],
        ["Số và ký hiệu văn bản", /^Số\s*[:.]?\s*\S+/im.test(text)],
        ["Địa danh, ngày tháng năm", /ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}/i.test(text)],
        ["Tên loại văn bản nhận diện được", lines.some(isTypeTitle)],
        ["Nơi nhận hoặc Kính gửi/Kính trình", /(^|\n)\s*(Nơi nhận|Kính gửi|Kính trình)\s*:/i.test(text)],
        ["Dòng chức vụ/quyền hạn ký", /(^|\n)\s*(T\/M\s|KT\.\s|PHÓ BÍ THƯ|BÍ THƯ|CHỦ TỊCH|GIÁM ĐỐC)/i.test(text)],
        ["Dấu cách thừa trong văn bản", !/(\S {2,}\S|\s+[,.!?;:])/.test(text)]
      ];
      const container = byId("review-results");
      container.replaceChildren(...checks.map(([label, found]) => {
        const row = document.createElement("div");
        row.className = "review-row";
        row.dataset.state = found ? "found" : "missing";
        const mark = document.createElement("span");
        mark.className = "mark";
        mark.textContent = found ? "✓" : "!";
        const caption = document.createElement("span");
        caption.textContent = label;
        row.append(mark, caption);
        return row;
      }));
      const issues = [];
      lines.forEach((line, index) => {
        if (!line) return;
        if (/\S {2,}\S|\s+[,.!?;:]/.test(line)) issues.push({ index, label: `Dòng ${index + 1}: có khoảng trắng thừa hoặc dấu câu đặt cách.` });
        if (/\bngày\s+\d{1,2}[/-]\d{1,2}[/-]\d{4}\b/i.test(line)) issues.push({ index, label: `Dòng ${index + 1}: ngày tháng đang viết dạng số; kiểm tra cách trình bày theo hướng dẫn.` });
      });
      if (issues.length) {
        const heading = document.createElement("p");
        heading.className = "hint";
        heading.textContent = `Phát hiện ${issues.length} vị trí cần xem:`;
        container.append(heading);
        issues.forEach(({ index, label }) => {
          const jump = document.createElement("button");
          jump.className = "review-row review-jump";
          jump.type = "button";
          jump.textContent = label;
          jump.addEventListener("click", () => selectParagraph(index).catch(showError));
          container.append(jump);
        });
      }
      setStatus(`Rà soát xong: ${issues.length} vị trí cần xem. Mục bắt buộc còn thiếu được đánh dấu bên trên.`, "success");
    });
  }

  async function selectParagraph(index) {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      if (index >= 0 && index < paragraphs.items.length) paragraphs.items[index].getRange().select();
      await context.sync();
    });
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) {
      setStatus("Add-in này chỉ hoạt động trong Word.", "error");
      return;
    }
    buttons.forEach((button) => { button.disabled = false; });
    byId("normalize").addEventListener("click", () => normalizeDocument().catch(showError));
    byId("title").addEventListener("click", () => formatSelection("title").catch(showError));
    byId("summary").addEventListener("click", () => formatSelection("summary").catch(showError));
    byId("review").addEventListener("click", () => reviewDocument().catch(showError));
    setStatus("Đã kết nối với Word. Hãy lưu một bản sao trước khi chuẩn hóa.", "success");
  });

  function showError(error) {
    console.error(error);
    setStatus(`Không thể hoàn tất thao tác: ${error && error.message ? error.message : "lỗi không xác định"}`, "error");
  }
})();

