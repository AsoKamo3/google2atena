document.getElementById("runBtn").addEventListener("click", async () => {
  const file = document.getElementById("fileInput").files[0];
  const status = document.getElementById("status");
  if (!file) { status.textContent = "CSVファイルを選択してください。"; return; }

  status.textContent = "変換中…";
  const fd = new FormData();
  fd.append("file", file);

  try {
    const res = await fetch("/convert", { method: "POST", body: fd });
    if (!res.ok) throw new Error("変換に失敗しました");
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "google_converted.csv";
    a.click();
    URL.revokeObjectURL(url);
    status.textContent = "✅ 変換が完了しました（google_converted.csv を保存）";
  } catch (e) {
    status.textContent = "⚠️ エラーが発生しました。CSVの形式や文字コードをご確認ください。";
  }
});
