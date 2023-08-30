const upload = document.querySelector(".form");
const fileInput = document.querySelector(".file-input");
const uploadedArea = document.querySelector(".uploaded-area");
const form = document.querySelector("form")

upload.addEventListener("click", () =>{
    fileInput.click();
  });
fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) {
    let fileName = file.name;
    if (fileName.length >= 12) {
      let splitName = fileName.split(".");
      fileName = splitName[0].substring(0, 13) + "... ." + splitName[1];
    }
    uploadFile(file, fileName);
  }
});

function uploadFile(file, name) {
    let uploadedHTML = `<div class="row">
                        <i class="fas fa-file-alt"></i>
                        <div class="content">
                            <div class="details">
                            <span class="name">${name} • Uploaded</span>
                            </div>
                        </div>
                        <i class="fas fa-check"></i>
                        </div>`;
    uploadedArea.innerHTML = uploadedHTML;
}