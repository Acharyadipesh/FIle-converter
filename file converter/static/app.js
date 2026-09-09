const tools =
    document.querySelectorAll(".tool");

const fileInput =
    document.getElementById("fileInput");

const dropzone =
    document.getElementById("dropzone");

const convertButton =
    document.getElementById("convertBtn");

const clearButton =
    document.getElementById("clearBtn");

const statusBox =
    document.getElementById("status");

const downloadButton =
    document.getElementById("download");

const fileInfo =
    document.getElementById("fileInfo");

const toolTitle =
    document.getElementById("toolTitle");

const toolDescription =
    document.getElementById("toolDesc");

const fileType =
    document.getElementById("fileType");

const compressionSettings =
    document.getElementById(
        "compressionSettings"
    );

const qualitySelect =
    document.getElementById(
        "qualitySelect"
    );


let currentTask =
    "pdf-to-ppt";

let currentAccept =
    ".pdf";

let currentMultiple =
    false;


const toolInformation = {

    "pdf-to-ppt": {

        title:
            "PDF to PowerPoint",

        description:
            "Upload a PDF and receive a PPTX presentation.",

        type:
            "PDF"

    },


    "ppt-to-pdf": {

        title:
            "PowerPoint to PDF",

        description:
            "Convert a PPT or PPTX presentation to PDF.",

        type:
            "PPT"

    },


    "jpg-to-png": {

        title:
            "JPG to PNG",

        description:
            "Convert JPEG images to PNG.",

        type:
            "JPG"

    },


    "png-to-jpg": {

        title:
            "PNG to JPG",

        description:
            "Convert PNG images to JPG.",

        type:
            "PNG"

    },


    "pdf-to-jpg": {

        title:
            "PDF to JPG",

        description:
            "Convert PDF pages into JPG images.",

        type:
            "PDF"

    },


    "jpg-to-pdf": {

        title:
            "JPG to PDF",

        description:
            "Create a PDF from JPG images.",

        type:
            "JPG"

    },


    "png-to-pdf": {

        title:
            "PNG to PDF",

        description:
            "Create a PDF from PNG images.",

        type:
            "PNG"

    },


    "merge-pdf": {

        title:
            "Merge PDF",

        description:
            "Combine multiple PDF files into one PDF.",

        type:
            "PDF"

    },


    "compress-image": {

        title:
            "Compress Image",

        description:
            "Reduce image file size while maintaining good quality.",

        type:
            "IMAGE"

    }

};


// ============================================================
// TOOL SELECTION
// ============================================================

tools.forEach(
    tool => {

        tool.addEventListener(
            "click",
            () => {

                tools.forEach(
                    item => {

                        item.classList.remove(
                            "active"
                        );

                    }
                );


                tool.classList.add(
                    "active"
                );


                currentTask =
                    tool.dataset.task;


                currentAccept =
                    tool.dataset.accept;


                currentMultiple =
                    tool.dataset.multiple ===
                    "true";


                const info =
                    toolInformation[
                        currentTask
                    ];


                toolTitle.textContent =
                    info.title;


                toolDescription.textContent =
                    info.description;


                fileType.textContent =
                    info.type;


                fileInput.accept =
                    currentAccept;


                if (
                    currentTask ===
                    "compress-image"
                ) {

                    fileInput.multiple =
                        false;

                    compressionSettings.hidden =
                        false;

                    convertButton.textContent =
                        "Compress image";

                }

                else {

                    fileInput.multiple =
                        currentMultiple ||
                        currentTask ===
                        "jpg-to-pdf" ||
                        currentTask ===
                        "png-to-pdf";

                    compressionSettings.hidden =
                        true;

                    convertButton.textContent =
                        "Convert now";

                }


                resetConverter();

            }
        );

    }
);


// ============================================================
// FILE INPUT
// ============================================================

fileInput.addEventListener(
    "change",
    () => {
        const files = fileInput.files;
        
        // Limit check for image-to-PDF conversions
        if (
            (currentTask === "jpg-to-pdf" || currentTask === "png-to-pdf") &&
            files.length > 5
        ) {
            showError("Maximum 5 images allowed for PDF conversion.");
            fileInput.value = "";
            fileInfo.textContent = "";
            return;
        }
        
        showSelectedFiles(files);
    }
);


// ============================================================
// SHOW FILES
// ============================================================

function showSelectedFiles(
    files
) {

    if (
        !files ||
        files.length === 0
    ) {

        fileInfo.textContent =
            "";

        return;

    }


    const names = [];


    for (
        const file of files
    ) {

        names.push(

            `${file.name} ` +
            `(${formatFileSize(file.size)})`

        );

    }


    fileInfo.textContent =
        names.join(" • ");

}


// ============================================================
// DRAG & DROP
// ============================================================

dropzone.addEventListener(
    "dragenter",
    event => {

        event.preventDefault();

        dropzone.classList.add(
            "drag"
        );

    }
);


dropzone.addEventListener(
    "dragover",
    event => {

        event.preventDefault();

        dropzone.classList.add(
            "drag"
        );

    }
);


dropzone.addEventListener(
    "dragleave",
    event => {

        event.preventDefault();

        dropzone.classList.remove(
            "drag"
        );

    }
);


dropzone.addEventListener(
    "drop",
    event => {

        event.preventDefault();

        dropzone.classList.remove(
            "drag"
        );

        const files =
            event.dataTransfer.files;
        
        // Limit check for image-to-PDF conversions
        if (
            (currentTask === "jpg-to-pdf" || currentTask === "png-to-pdf") &&
            files.length > 5
        ) {
            showError("Maximum 5 images allowed for PDF conversion.");
            return;
        }

        if (
            currentTask === "compress-image" &&
            files.length > 1
        ) {

            showError(
                "Image compression accepts only one image."
            );

            return;

        }


        fileInput.files =
            files;


        showSelectedFiles(
            files
        );

    }
);


// ============================================================
// CONVERT
// ============================================================

convertButton.addEventListener(
    "click",
    async () => {

        if (
            !fileInput.files ||
            fileInput.files.length === 0
        ) {

            showError(
                "Please choose a file first."
            );

            return;

        }


        if (

            currentTask ===
            "compress-image"

            &&

            fileInput.files.length !== 1

        ) {

            showError(
                "Please select exactly one image."
            );

            return;

        }


        convertButton.disabled =
            true;


        downloadButton.hidden =
            true;


        setStatus(

            currentTask ===
            "compress-image"

                ?

                "Compressing image..."

                :

                "Converting..."

        );


        const formData =
            new FormData();


        formData.append(
            "task",
            currentTask
        );


        for (
            const file of fileInput.files
        ) {

            formData.append(
                "files",
                file
            );

        }


        if (
            currentTask ===
            "compress-image"
        ) {

            formData.append(

                "quality",

                qualitySelect.value

            );

        }


        try {

            const response =
                await fetch(

                    "/api/convert",

                    {

                        method:
                            "POST",

                        body:
                            formData

                    }

                );


            let data;


            try {

                data =
                    await response.json();

            }

            catch {

                throw new Error(
                    "The server returned an invalid response."
                );

            }


            if (!response.ok) {

                throw new Error(

                    data.error ||
                    "Conversion failed."

                );

            }


            if (
                currentTask ===
                "compress-image"
            ) {

                const original =
                    formatFileSize(
                        data.original_size
                    );


                const output =
                    formatFileSize(
                        data.output_size
                    );


                let saved =
                    data.saved_percent;


                if (saved < 0) {

                    saved = 0;

                }


                setStatus(

                    `Compression complete! ` +

                    `Original: ${original} → ` +

                    `Compressed: ${output} ` +

                    `(${saved}% smaller)`

                );

            }

            else {

                setStatus(
                    "Conversion complete!"
                );

            }


            downloadButton.href =
                data.url;


            downloadButton.textContent =
                `Download ${data.filename}`;


            downloadButton.hidden =
                false;

        }


        catch (error) {

            console.error(
                error
            );


            showError(
                error.message
            );

        }


        finally {

            convertButton.disabled =
                false;

        }

    }
);


// ============================================================
// CLEAR
// ============================================================

clearButton.addEventListener(
    "click",
    () => {

        resetConverter();

    }
);


// ============================================================
// RESET
// ============================================================

function resetConverter() {

    fileInput.value =
        "";

    fileInfo.textContent =
        "";

    statusBox.textContent =
        "";

    statusBox.className =
        "status";

    downloadButton.hidden =
        true;

}


// ============================================================
// STATUS
// ============================================================

function setStatus(
    message
) {

    statusBox.textContent =
        message;

    statusBox.className =
        "status";

}


function showError(
    message
) {

    statusBox.textContent =
        message;

    statusBox.className =
        "status error";

}


// ============================================================
// FILE SIZE
// ============================================================

function formatFileSize(
    bytes
) {

    if (
        !bytes ||
        bytes <= 0
    ) {

        return "0 Bytes";

    }


    const units = [

        "Bytes",

        "KB",

        "MB",

        "GB"

    ];


    const index =
        Math.floor(

            Math.log(bytes)
            /
            Math.log(1024)

        );


    return (

        parseFloat(

            (

                bytes /
                Math.pow(
                    1024,
                    index
                )

            ).toFixed(2)

        )

        +

        " "

        +

        units[index]

    );

}


// ============================================================
// THEME TOGGLE — DARK / LIGHT MODE
// ============================================================

const themeToggle =
    document.getElementById("themeToggle");

const body =
    document.body;


// Check saved preference from localStorage
const savedTheme =
    localStorage.getItem("theme");

if (savedTheme === "dark") {

    body.classList.add("dark");

    themeToggle.textContent = "☀️";

} else {

    themeToggle.textContent = "🌙";

}


// Toggle theme on button click
themeToggle.addEventListener(
    "click",
    function() {

        body.classList.toggle("dark");

        const isDark =
            body.classList.contains("dark");

        // Update button icon
        this.textContent =
            isDark ? "☀️" : "🌙";

        // Save preference
        localStorage.setItem(
            "theme",
            isDark ? "dark" : "light"
        );

    }
);


// ============================================================

// ============================================================

const yearElement = document.getElementById('year');
if (yearElement) {
    yearElement.textContent = new Date().getFullYear();
}


// ============================================================

// ============================================================

document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
        const targetId = this.getAttribute('href');
        const targetElement = document.querySelector(targetId);

        if (targetElement) {
            e.preventDefault();
            const headerOffset = 80; // Height of fixed header
            const elementPosition = targetElement.getBoundingClientRect().top;
            const offsetPosition = elementPosition + window.pageYOffset - headerOffset;

            window.scrollTo({
                top: offsetPosition,
                behavior: 'smooth'
            });
        }
    });
});
