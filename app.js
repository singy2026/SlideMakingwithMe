/* SlideMakingwithMe — UI wiring */

const mediaType = document.getElementById("media-type");
const pastorField = document.getElementById("pastor-field");
const pastor = document.getElementById("pastor");
const fileField = document.getElementById("file-field");
const dropZone = document.getElementById("drop-zone");
const dropText = document.getElementById("drop-text");
const fileInput = document.getElementById("docx-file");
const convertBtn = document.getElementById("convert-btn");
const unsupportedNote = document.getElementById("unsupported-note");
const errorMsg = document.getElementById("error-msg");
const resultCard = document.getElementById("result-card");
const resultTitle = document.getElementById("result-title");
const resultSub = document.getElementById("result-sub");
const preview = document.getElementById("preview");
const downloadBtn = document.getElementById("download-btn");

let selectedFile = null;
let lastResult = null; // { name, pro }

const SUPPORTED = { media: "sermons", pastor: "jon-choi" };

function refresh() {
  const media = mediaType.value;
  const isSermons = media === "sermons";
  pastorField.classList.toggle("hidden", !isSermons);

  let supported = false;
  let noteText = "";
  if (media === "songs") {
    noteText = "🎵 Songs conversion is coming soon. For now, please use Sermons.";
  } else if (isSermons && pastor.value && pastor.value !== SUPPORTED.pastor) {
    noteText = "⏳ This pastor's template is coming soon. Currently supported: P Jon Choi.";
  } else if (isSermons && pastor.value === SUPPORTED.pastor) {
    supported = true;
  }

  unsupportedNote.textContent = noteText;
  unsupportedNote.classList.toggle("hidden", !noteText);
  fileField.classList.toggle("hidden", !supported);
  convertBtn.classList.toggle("hidden", !supported);
  convertBtn.disabled = !(supported && selectedFile);
  errorMsg.classList.add("hidden");
}

mediaType.addEventListener("change", refresh);
pastor.addEventListener("change", refresh);

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) setFile(fileInput.files[0]);
});

function setFile(file) {
  if (!file.name.toLowerCase().endsWith(".docx")) {
    showError("Please choose a Word document (.docx).");
    return;
  }
  selectedFile = file;
  dropZone.classList.add("has-file");
  dropText.innerHTML = "📄 <strong></strong>";
  dropText.querySelector("strong").textContent = file.name;
  refresh();
}

function showError(msg) {
  errorMsg.textContent = msg;
  errorMsg.classList.remove("hidden");
}

convertBtn.addEventListener("click", async () => {
  if (!selectedFile) return;
  convertBtn.disabled = true;
  convertBtn.textContent = "Converting…";
  errorMsg.classList.add("hidden");
  try {
    const buf = await selectedFile.arrayBuffer();
    const name = selectedFile.name.replace(/\.docx$/i, "");
    const { slides, pro } = await convertSermonDocx(buf, name, TEMPLATE_JONCHOI_B64);
    lastResult = { name, pro };
    renderResult(name, slides);
  } catch (err) {
    showError("Conversion failed: " + err.message);
  } finally {
    convertBtn.disabled = false;
    convertBtn.textContent = "Convert to ProPresenter";
  }
});

function renderResult(name, slides) {
  resultTitle.textContent = name + ".pro";
  resultSub.textContent = slides.length + " slides · Sermon · P Jon Choi template";
  preview.innerHTML = "";
  slides.forEach((text, i) => {
    const div = document.createElement("div");
    div.className = "slide";
    const num = document.createElement("span");
    num.className = "num";
    num.textContent = "Slide " + (i + 1);
    const p = document.createElement("p");
    p.textContent = text;
    div.append(num, p);
    preview.appendChild(div);
  });
  resultCard.classList.remove("hidden");
  resultCard.scrollIntoView({ behavior: "smooth" });
}

downloadBtn.addEventListener("click", () => {
  if (!lastResult) return;
  const blob = new Blob([lastResult.pro], { type: "application/octet-stream" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = lastResult.name + ".pro";
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
});
