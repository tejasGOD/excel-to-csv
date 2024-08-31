document.addEventListener("DOMContentLoaded", () => {
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("fileInput");
  const convertBtn = document.getElementById("convertBtn");
  const output = document.getElementById("output");

  dropZone.addEventListener("click", () => {
    fileInput.click();
  });

  fileInput.addEventListener("change", handleFileSelect);

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    const files = e.dataTransfer.files;
    handleFiles(files);
  });

  function handleFileSelect(e) {
    const files = e.target.files;
    handleFiles(files);
  }

  function handleFiles(files) {
    if (files.length > 0) {
      const file = files[0];
      if (
        file.type === "application/vnd.ms-excel" ||
        file.type ===
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ) {
        output.textContent = `Selected file: ${file.name}`;
        convertBtn.disabled = false;
      } else {
        output.textContent = "Please upload a valid Excel document.";
        convertBtn.disabled = true;
      }
    }
  }

  convertBtn.addEventListener("click", () => {
    const file = fileInput.files[0];
    if (file) {
      output.textContent = "Converting...";
      convertExcelToCSV(file);
    } else {
      output.textContent = "No file selected.";
    }
  });

  function convertExcelToCSV(file) {
    const formData = new FormData();
    formData.append("file", file);

    fetch(
      "https://api.convertapi.com/convert/xlsx/to/csv?Secret=YOUR_API_SECRET",
      {
        method: "POST",
        body: formData,
      }
    )
      .then((response) => response.json())
      .then((data) => {
        if (data.Files && data.Files.length > 0) {
          const csvUrl = data.Files[0].Url;
          output.innerHTML = `<a href="${csvUrl}" target="_blank">Download CSV</a>`;
        } else {
          output.textContent = "Conversion failed. Please try again.";
        }
      })
      .catch((error) => {
        console.error("Error:", error);
        output.textContent = "Conversion failed. Please try again.";
      });
  }

  ScrollReveal().reveal(".container", { delay: 200 });
});
