function onEdit(e) {
    const sheet = e.range.getSheet();

    // Change to your sheet name
    if (sheet.getName() !== "Manager Sheet") return;

    // Only run for column G and rows 12 onwards
    if (e.range.getColumn() !== 7 || e.range.getRow() < 12) return;

    const id = String(e.value || "").trim();
    if (!id) return;

    const url = `https://docs.google.com/spreadsheets/d/${id}`;

    const richText = SpreadsheetApp.newRichTextValue()
        .setText(id)
        .setLinkUrl(url)
        .build();

    e.range.setRichTextValue(richText);
}