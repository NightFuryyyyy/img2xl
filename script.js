import ExcelJS from 'https://cdn.skypack.dev/exceljs';

function pixelToHex(pixel) {
    return (
        pixel[3].toString(16).padStart(2, "0") +
        pixel[0].toString(16).padStart(2, "0") +
        pixel[1].toString(16).padStart(2, "0") +
        pixel[2].toString(16).padStart(2, "0")
    ).toUpperCase();
}

const removeExtension = file_name => file_name.slice(0, file_name.lastIndexOf("."));

const image_input = document.querySelector(".image_input");
const preview_canvas = document.querySelector(".preview_canvas");
const export_button = document.querySelector(".export_button");

const canvas_context = preview_canvas.getContext("2d");

image_input.onchange = () => {
    const image_file = image_input.files[0];
    if(!image_file) {
        alert('Please upload an image.');
        return;
    }

    const reader = new FileReader();

    reader.onload = event => {
        const image_object = new Image();

        image_object.onload = async function() {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Sheet1');
            worksheet.properties.defaultColWidth = 2.9;

            let canvas_height = image_object.height;
            let canvas_width = image_object.width;
            if (canvas_width > 854) {
                canvas_width = 854;
                canvas_height = Math.floor((canvas_width / image_object.width) * image_object.height);
            }
            if (canvas_height > 480) {
                canvas_height = 480;
                canvas_width = Math.floor((canvas_height / image_object.height) * image_object.width);
            }
            preview_canvas.height = canvas_height;
            preview_canvas.width = canvas_width;
            canvas_context.drawImage(image_object, 0, 0, canvas_width, canvas_height);
            for (let i = 0; i < canvas_height; i++) {
                for (let j = 0; j < canvas_width; j++) {
                    const pixel = canvas_context.getImageData(j, i, 1, 1).data;
                    const cell = worksheet.getCell(i + 1, j + 1);
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: pixelToHex(pixel) }
                    };
                }
            }
            const final_cell = worksheet.getCell(canvas_height, canvas_width);
            const final_pixel = canvas_context.getImageData(canvas_width - 1, canvas_height - 1, 1, 1).data;
            final_cell.value = 1;
            final_cell.font = {
                color: { argb: pixelToHex(final_pixel) }
            };

            export_button.onclick = async function() {
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const a = document.createElement('a');
                a.href = URL.createObjectURL(blob);
                a.download = `${removeExtension(image_file.name)}.xlsx`
                a.click();
            }
        };

        image_object.src = event.target.result;
    };

    reader.readAsDataURL(image_file);
}