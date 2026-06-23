const fileInput = document.querySelector("#fileInput");
const convertBtn = document.querySelector("#convertBtn");
const dropzone = document.querySelector("#dropzone");
const fileList = document.querySelector("#fileList");
const statusEl = document.querySelector("#status");
const resultsEl = document.querySelector("#results");
const customerSlotsEl = document.querySelector("#customerSlots");
const refreshCustomersBtn = document.querySelector("#refreshCustomersBtn");

const MAX_UPLOADS = 10;
let selectedFiles = [];
let customerSlots = [];
let downloadObjectUrls = [];

function fileToBase64(file) {
  return file.arrayBuffer().then((buffer) => {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
    return btoa(binary);
  });
}

function setStatus(text) {
  statusEl.textContent = text;
}

function clearDownloadObjectUrls() {
  for (const url of downloadObjectUrls) URL.revokeObjectURL(url);
  downloadObjectUrls = [];
}

function base64ToBlob(base64, mimeType) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mimeType });
}

function renderFiles() {
  fileList.innerHTML = "";
  for (const file of selectedFiles) {
    const li = document.createElement("li");
    li.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
    fileList.appendChild(li);
  }
  setStatus(selectedFiles.length ? `${selectedFiles.length} ไฟล์พร้อมแปลง` : "รอเลือกไฟล์");
}

function setFiles(files) {
  selectedFiles = Array.from(files || []).slice(0, MAX_UPLOADS);
  renderFiles();
}

function selectedCustomerSlotIds() {
  return customerSlots.filter((slot) => slot.selected && slot.nameCount > 0).map((slot) => slot.id);
}

function customerSummary(slot) {
  const label = slot.label || slot.originalFileName || slot.fileName || slot.id;
  return `${label} · ${slot.nameCount || 0} unique usable names`;
}

function renderCustomerSlots() {
  customerSlotsEl.innerHTML = "";
  if (!customerSlots.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "ยังไม่มี customer list ที่บันทึกไว้";
    customerSlotsEl.appendChild(empty);
    return;
  }

  for (const slot of customerSlots) {
    const card = document.createElement("div");
    card.className = "slot";

    const top = document.createElement("div");
    top.className = "slot-top";

    const left = document.createElement("label");
    left.className = "slot-select";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = Boolean(slot.selected);
    checkbox.addEventListener("change", () => {
      slot.selected = checkbox.checked;
    });
    const labelText = document.createElement("span");
    labelText.textContent = customerSummary(slot);
    left.append(checkbox, labelText);

    const actions = document.createElement("div");
    actions.className = "slot-actions";

    const upload = document.createElement("label");
    upload.className = "ghost-button";
    upload.textContent = "อัปโหลด/แทนที่";
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xls,.csv";
    input.hidden = true;
    input.addEventListener("change", async () => {
      if (!input.files[0]) return;
      await saveCustomerSlot(slot.id, input.files[0]);
      input.value = "";
    });
    upload.appendChild(input);

    const del = document.createElement("button");
    del.type = "button";
    del.className = "danger-button";
    del.textContent = "ลบ";
    del.addEventListener("click", async () => {
      await deleteCustomerSlot(slot.id);
    });

    actions.append(upload, del);
    top.append(left, actions);

    const meta = document.createElement("div");
    meta.className = "slot-meta";
    meta.textContent = slot.fileName
      ? `ไฟล์: ${slot.fileName}${slot.detectedColumnLabel ? ` · อ่านจากคอลัมน์ ${slot.detectedColumnLabel}` : ""}${slot.usableCount ? ` · พบชื่อที่ใช้ได้ ${slot.usableCount} แถว` : ""}`
      : "ยังไม่ได้อัปโหลดไฟล์";

    card.append(top, meta);
    customerSlotsEl.appendChild(card);
  }
}

async function loadCustomerSlots() {
  const response = await fetch("/api/customer-lists");
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "โหลด customer list ไม่สำเร็จ");

  const existingSelection = new Map(customerSlots.map((slot) => [slot.id, Boolean(slot.selected)]));
  customerSlots = Array.from({ length: data.maxSlots || 5 }, (_, index) => {
    const id = `slot-${index + 1}`;
    const stored = (data.slots || []).find((slot) => slot.id === id);
    return {
      id,
      label: stored?.label || `Customer list ${index + 1}`,
      fileName: stored?.fileName || "",
      originalFileName: stored?.originalFileName || "",
      nameCount: stored?.nameCount || 0,
      usableCount: stored?.usableCount || 0,
      detectedColumnLabel: stored?.detectedColumnLabel || "",
      selected: stored ? existingSelection.get(id) ?? true : false,
    };
  });
  renderCustomerSlots();
}

async function saveCustomerSlot(slotId, file) {
  setStatus(`กำลังบันทึก ${file.name}...`);
  const response = await fetch("/api/customer-lists", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      slotId,
      file: { name: file.name, base64: await fileToBase64(file) },
    }),
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "บันทึก customer list ไม่สำเร็จ");
  await loadCustomerSlots();
  setStatus(`บันทึก customer list เรียบร้อย: ${data.slot.fileName}`);
}

async function deleteCustomerSlot(slotId) {
  setStatus("กำลังลบ customer list...");
  const response = await fetch(`/api/customer-lists/${encodeURIComponent(slotId)}`, { method: "DELETE" });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "ลบ customer list ไม่สำเร็จ");
  await loadCustomerSlots();
  setStatus("ลบ customer list เรียบร้อย");
}

function preventWindowDrop(event) {
  event.preventDefault();
}

["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
  window.addEventListener(eventName, preventWindowDrop);
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragover");
  setFiles(event.dataTransfer.files);
});

dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => setFiles(fileInput.files));

refreshCustomersBtn.addEventListener("click", async () => {
  refreshCustomersBtn.disabled = true;
  try {
    await loadCustomerSlots();
  } catch (error) {
    setStatus(`เกิดข้อผิดพลาด: ${error.message}`);
  } finally {
    refreshCustomersBtn.disabled = false;
  }
});

convertBtn.addEventListener("click", async () => {
  clearDownloadObjectUrls();
  resultsEl.innerHTML = "";

  if (!selectedFiles.length) {
    setStatus("กรุณาเลือกไฟล์ก่อน");
    return;
  }

  const customerSlotIds = selectedCustomerSlotIds();
  if (!customerSlotIds.length) {
    const message = "ยังไม่ได้เลือก customer list จึงแปลงไฟล์ไม่ได้";
    setStatus(message);
    window.alert(message);
    return;
  }

  convertBtn.disabled = true;
  setStatus("กำลังแปลงไฟล์...");

  try {
    const files = [];
    for (const file of selectedFiles) {
      files.push({ name: file.name, base64: await fileToBase64(file) });
    }

    const response = await fetch("/api/convert", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ files, customerSlotIds }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Conversion failed");

    setStatus(`แปลงเสร็จ ${data.results.length} ไฟล์`);
    for (const item of data.results) {
      const row = document.createElement("div");
      row.className = "result";
      const blob = base64ToBlob(
        item.outputBase64,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      const objectUrl = URL.createObjectURL(blob);
      downloadObjectUrls.push(objectUrl);
      row.innerHTML = `<div>${item.outputName}</div><a href="${objectUrl}" download="${item.outputName}">ดาวน์โหลด</a>`;
      resultsEl.appendChild(row);
    }
  } catch (error) {
    setStatus(`เกิดข้อผิดพลาด: ${error.message}`);
  } finally {
    convertBtn.disabled = false;
  }
});

renderFiles();
loadCustomerSlots().catch((error) => setStatus(`โหลด customer list ไม่สำเร็จ: ${error.message}`));
