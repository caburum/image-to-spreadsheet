const Jimp = require('jimp');
const ExcelJS = require('exceljs/dist/es5');

const replitRun = process.argv.indexOf('index.js'); // Replit adds an extra arg, this fixes the path
const imagePath = (replitRun != -1 ? process.argv[replitRun + 1] : process.argv[2]) || 'example.png';

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Image', {
	views: [{
		showGridLines: false
	}],
	properties: { // This is a bit broken
		defaultRowHeight: 90,
		defaultColWidth: 30
	}
});

Jimp.read(imagePath)
	.then(async (imageData) => { // Convert image into array (row => column => pixel values)
		imageData = await imageData.scaleToFit(85, 65000); // Max of 256 columns / 3
		var image = Array.from(Array(imageData.bitmap.height), () => new Array(imageData.bitmap.width));
		for (var y = 0; y < imageData.bitmap.height; y++) { // Each row
			for (var x = 0; x < imageData.bitmap.width; x++) { // Each column
				let pixel = imageData.getPixelColor(x, y);
				let { r, g, b, a } = Jimp.intToRGBA(pixel);
				image[y][x] = [r, g, b];
			}
		}
		return image;
	})
	.then((image) => {
		image.forEach((column, rowIndex) => { // Each column in row
			column.forEach((pixel, colIndex) => { // Each pixel in column
				pixel.forEach((pixelValue, pixelIndex) => { // Each 3 colors in pixel
					let cell = sheet.getCell(rowIndex, ((colIndex) * 3) + (pixelIndex + 1));

					cell.value = pixelValue;

					let color = ['00', '00', '00'];
					color[pixelIndex] = pixelValue.toString(16);
					cell.fill = {
						type: 'pattern',
						pattern: 'solid',
						fgColor: {
							argb: 'FF' + color.join('')
						}
					}
				});
			});
		})
		workbook.xlsx.writeFile('output.xlsx');
	});