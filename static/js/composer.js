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
  const editor = document.getElementById("detailsEditor");
  const form = document.getElementById("composeForm");
  const syncDetails = () => { details.value = editor.innerText.trim(); };
  const setEditorText = text => {
    editor.textContent = text;
    syncDetails();
  };
  const appendEditorText = text => {
    setEditorText([editor.innerText.trim(), text.trim()].filter(Boolean).join("\n"));
  };

  document.querySelectorAll("[data-command]").forEach(button => {
    button.addEventListener("click", () => {
      editor.focus();
      document.execCommand(button.dataset.command, false, null);
      syncDetails();
    });
  });
  editor.addEventListener("input", syncDetails);
  form.addEventListener("submit", event => {
    syncDetails();
    if (!details.value) {
      event.preventDefault();
      editor.focus();
      editor.classList.add("is-invalid");
    }
  });
  editor.addEventListener("input", () => editor.classList.remove("is-invalid"));

  const subject = document.getElementById("subject");
  const mainSubject = document.getElementById("mainSubject");
  const suggestionStatus = document.getElementById("suggestionStatus");
  let suggestionRequest = 0;
  const requestSuggestion = async () => {
    const requestNumber = ++suggestionRequest;
    suggestionStatus.textContent = "Creating a suggestion…";
    try {
      const response = await fetch("/generate-predefined-content", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
          area: group.value,
          category: category.value,
          start_date: document.getElementById("startDate").value,
          end_date: document.getElementById("endDate").value,
          main_subject: mainSubject.value,
          language: language.value
        })
      });
      const result = await response.json();
      if (!response.ok || !result.success) throw new Error(result.error || "Unable to create a suggestion.");
      if (requestNumber !== suggestionRequest) return;
      subject.value = result.title;
      setEditorText(result.description);
      suggestionStatus.textContent = "Suggested title and description added. You can edit both before generating.";
    } catch (error) {
      suggestionStatus.textContent = error.message;
    }
  };
  document.getElementById("suggestContent").addEventListener("click", requestSuggestion);
  let suggestionTimer;
  [mainSubject, document.getElementById("startDate"), document.getElementById("endDate")]
    .forEach(control => control.addEventListener("input", () => {
      window.clearTimeout(suggestionTimer);
      suggestionTimer = window.setTimeout(requestSuggestion, 650);
    }));
  category.addEventListener("change", requestSuggestion);
  group.addEventListener("change", requestSuggestion);
  const applyWritingDirection = () => {
    const isRtl = ["ar", "ur"].includes(language.value);
    subject.dir = isRtl ? "rtl" : "ltr";
    editor.dir = isRtl ? "rtl" : "ltr";
  };
  language.addEventListener("change", () => {
    applyWritingDirection();
    requestSuggestion();
  });
  applyWritingDirection();
  requestSuggestion();
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
      if (finalText) appendEditorText(finalText);
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
      appendEditorText(result.text);
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
