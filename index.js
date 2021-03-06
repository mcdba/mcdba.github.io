document.addEventListener("load", main);


function NumToCol(num) {
	for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
		ret = String.fromCharCode(parseInt((num % b) / a) + 65) + ret;
	}
	return ret;
}
function ColToNum(val) {
	var base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', i, j, result = 0;
	for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
		result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
	}
	return result;
};
function update_sheet_range(ws) {
	var range = { s: { r: Infinity, c: Infinity }, e: { r: 0, c: 0 } };
	Object.keys(ws).filter(function (x) { return x.charAt(0) != "!"; }).map(XLSX.utils.decode_cell).forEach(function (x) {
		range.s.c = Math.min(range.s.c, x.c); range.s.r = Math.min(range.s.r, x.r);
		range.e.c = Math.max(range.e.c, x.c); range.e.r = Math.max(range.e.r, x.r);
	});
	ws['!ref'] = XLSX.utils.encode_range(range);
}
function handle_workbook(workbook, resolve) {
	let result = [];
	for (let i = 0; i < workbook.SheetNames.length; i++) {
		let sheet = workbook.Sheets[workbook.SheetNames[i]];
		update_sheet_range(sheet);
		let range = sheet['!ref']
		if (!range) {
			throw new Error(workbook.SheetNames[i] + " is empty");
		}
		let rangeSplit = range.split(':');
		let rangeFrom = rangeSplit[0];
		let rangeTo = rangeSplit[1];
		let Col_Reg = RegExp("[A-Z]+");
		let colFrom = Col_Reg.exec(rangeFrom)[0];
		let colTo = Col_Reg.exec(rangeTo)[0];

		let Row_Reg = RegExp("[0-9]+");
		let rowFrom = Row_Reg.exec(rangeFrom)[0];
		let rowTo = Row_Reg.exec(rangeTo)[0];

		let table = [];
		for (let i = parseInt(rowFrom); i <= parseInt(rowTo); i++) {
			let row = {};
			for (let j = ColToNum(colFrom); j <= ColToNum(colTo); j++) {
				let index = NumToCol(j) + i.toString();
				let data = sheet[index];
				if (data) {
					row[j] = data.v;
				}
			}
			table.push(row);
		}
		result.push(table);
	}
	resolve(result);
}

function outWorkbook(data) {
	console.log('Out data...')
	const sheet = data[2]

	const columnRow = sheet[1]
	const columns = {}
	for (var i = 3; i <= 36; i += 1) {
		if (i <= 15 || i >= 19) {
			columns[i] = columnRow[i].split(' ')[0]
		}
	}
	console.log(columns)

	const reportData = []
	// step by date
	for (i = 0; i < 1395; i += 45) {
		let date = `${sheet[i][27]}-${sheet[i][28]}-${sheet[i][32]}`
		// console.log(dateStr)
		let data = []
		for (row = i + 2; row < i + 42; row += 1) {
			const dataRow = sheet[row]
			const name = dataRow[2]
			const uslugi = []
			for (var key in dataRow) {
				if ((key >= 3 && key <= 15) || (key >= 19 && key <= 36)) {
					if (dataRow[key] === 1) {
						uslugi.push(columns[key])
					}
				}
			}
			if (uslugi.length > 0) {
				data.push({ name, uslugi })
			}
		}
		if (data.length > 0) {
			reportData.push({ date, data })
		}

	}
	document.querySelector('#workbookContainer').innerHTML = toHTML(reportData)
	console.log(sheet)
	console.log(reportData)
}
function uslugiToHTML(data) {
	let html = ''
	data.forEach(u => {
		html = html + `<span>${u}</span>`
	})
	return html
}
function personsToHTML(data) {
	let html = ''
	data.forEach(u => {
		html = html + `
		<div style='display: flex'>
			<div class='name' style='width:10em'>${u.name}</div>
			<div class='uslugi'>${uslugiToHTML(u.uslugi)}</div>
		</div>
		`
	})
	return html
}
function toHTML(data) {
	let html = ''
	data.forEach(d => {
		html = html + `
		<div class='data'>
			<h3>${d.date}</h3>
			${personsToHTML(d.data)}
		</div>
		`
	})
	return html
}
function onFileEnter(e) {
	var inputFile = e.target.files[0];
	var fileReader = new FileReader();

	fileReader.onload = function (e) {
		var data = e.target.result;
		var workbook = XLSX.read(data, {
			type: 'binary'
		});
		handle_workbook(workbook, outWorkbook)
	};
	fileReader.readAsBinaryString(inputFile);
}


function main() {
	const inputFile = document.querySelector('#inputFile')
	console.log('dom loaded')
	inputFile.addEventListener('change', onFileEnter)
}
