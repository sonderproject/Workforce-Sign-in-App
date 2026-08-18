// Prevent double-submit on kiosk forms (guards against duplicate intake/visit).
document.addEventListener("submit", function (e) {
  const form = e.target;
  if (!form.matches("[data-guard]")) return;
  const btn = form.querySelector("button[type=submit]");
  if (btn) {
    if (btn.dataset.submitting === "1") { e.preventDefault(); return; }
    btn.dataset.submitting = "1";
    setTimeout(function () { btn.disabled = true; btn.textContent = "Please wait…"; }, 0);
  }
});

// Drag-and-drop workbook uploader (admin dashboard).
(function () {
  const zone = document.getElementById("dropzone");
  if (!zone) return;
  const input = zone.querySelector('input[type=file]');
  const nameOut = zone.querySelector(".filename");
  const form = zone.closest("form");

  function show(file) {
    if (file && nameOut) {
      nameOut.textContent = "Selected: " + file.name;
      if (form) form.querySelector("[data-upload-btn]").disabled = false;
    }
  }
  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => show(input.files[0]));
  ["dragenter", "dragover"].forEach(ev =>
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add("drag"); }));
  ["dragleave", "drop"].forEach(ev =>
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove("drag"); }));
  zone.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; show(input.files[0]); }
  });
})();
