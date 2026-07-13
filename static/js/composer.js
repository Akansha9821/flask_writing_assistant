document.addEventListener("DOMContentLoaded", () => {
  const tree = JSON.parse(document.getElementById("categoryTreeData").textContent);
  const group = document.getElementById("categoryGroup");
  const category = document.getElementById("category");

  function loadGroups() {
    Object.keys(tree).forEach(name => group.add(new Option(name, name)));
    loadCategories();
  }
  function loadCategories() {
    category.innerHTML = "";
    Object.entries(tree[group.value] || {}).forEach(([value, label]) => {
      category.add(new Option(label, value));
    });
  }
  group.addEventListener("change", loadCategories);
  loadGroups();

  const language = document.getElementById("language");
  const localeMap = {hi:"hi",mr:"mr",bn:"bn",gu:"gu",pa:"pa",ta:"ta",te:"te",kn:"kn",ml:"ml",ur:"ur",ne:"ne",fr:"fr",de:"de",es:"es",it:"it",pt:"pt",ar:"ar",ja:"ja",ko:"ko",ru:"ru",zh:"zh-CN"};
  const prefix = (navigator.language || "en").toLowerCase().split("-")[0];
  if (localeMap[prefix]) language.value = localeMap[prefix];

  const details = document.getElementById("details");
  const voiceButton = document.getElementById("voiceButton");
  const voiceStatus = document.getElementById("voiceStatus");
  const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  let recognition;
  if (Recognition) {
    recognition = new Recognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.onstart = () => {
      voiceButton.classList.add("btn-danger");
      voiceStatus.textContent = "Listening… click the button again to stop.";
    };
    recognition.onresult = event => {
      let finalText = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        if (event.results[i].isFinal) finalText += event.results[i][0].transcript + " ";
      }
      if (finalText) details.value += (details.value ? " " : "") + finalText;
    };
    recognition.onend = () => {
      voiceButton.classList.remove("btn-danger");
      voiceStatus.textContent = "Voice input stopped.";
    };
    voiceButton.addEventListener("click", () => {
      recognition.lang = language.value === "hi" ? "hi-IN" : navigator.language;
      try { recognition.start(); } catch (_) { recognition.stop(); }
    });
  } else {
    voiceButton.disabled = true;
    voiceStatus.textContent = "Speech recognition is unavailable in this browser.";
  }

  document.getElementById("extractText").addEventListener("click", async () => {
    const file = document.getElementById("handwritingImage").files[0];
    const status = document.getElementById("ocrStatus");
    if (!file) { status.textContent = "Select an image first."; return; }
    const data = new FormData();
    data.append("handwriting_image", file);
    data.append("ocr_language", document.getElementById("ocrLanguage").value);
    status.textContent = "Extracting handwritten text…";
    try {
      const response = await fetch("/ocr", {method:"POST", body:data});
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      details.value += (details.value ? "\n" : "") + result.text;
      status.textContent = "Text added. Review and correct it before generating.";
    } catch (error) {
      status.textContent = error.message;
    }
  });

  const canvas = document.getElementById("signatureCanvas");
  const ctx = canvas.getContext("2d");
  const hidden = document.getElementById("signatureData");
  let drawing = false;
  function p(event) {
    const rect = canvas.getBoundingClientRect();
    const source = event.touches ? event.touches[0] : event;
    return {x:(source.clientX-rect.left)*(canvas.width/rect.width), y:(source.clientY-rect.top)*(canvas.height/rect.height)};
  }
  function start(event){ drawing=true; const point=p(event); ctx.beginPath(); ctx.moveTo(point.x,point.y); event.preventDefault(); }
  function draw(event){ if(!drawing)return; const point=p(event); ctx.lineWidth=3; ctx.lineCap="round"; ctx.strokeStyle="#111"; ctx.lineTo(point.x,point.y); ctx.stroke(); event.preventDefault(); }
  function stop(){ if(!drawing)return; drawing=false; hidden.value=canvas.toDataURL("image/png"); }
  ["mousedown","touchstart"].forEach(e=>canvas.addEventListener(e,start));
  ["mousemove","touchmove"].forEach(e=>canvas.addEventListener(e,draw));
  ["mouseup","mouseleave","touchend"].forEach(e=>canvas.addEventListener(e,stop));
  document.getElementById("clearSignature").addEventListener("click",()=>{ctx.clearRect(0,0,canvas.width,canvas.height);hidden.value="";});
});
