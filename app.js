function toJSON(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function (sheetName) {
		var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if (roa.length > 0) {
			result[sheetName] = roa;
		}
	});
	return result;
};

function displayJSON(json) {

	var tabs = $("<ul/>").addClass("nav nav-tabs");
	var tabContents = $("<div/>").addClass("tab-content");
	var first = true;
	for (prop in json) {

		var tab = $("<li/>").append($("<a/>").attr("href", "#" + prop).attr("data-toggle", "tab").html(prop));
		var tabContent = $("<div/>").addClass("tab-pane").attr("id", prop).append($("<table/>").attr("id", prop + "_table").addClass("table table-stripped table-hover"));
		if (first) {
			tab.addClass("active");
			tabContent.addClass("active");
			first = false;
		}
		tabs.append(tab);
		tabContents.append(tabContent);
	}

	$("#workbookContainer").empty();
	$("#workbookContainer").append(tabs);
	$("#workbookContainer").append(tabContents);
	for (prop in json) {
		var columns = [];
		for (column in json[prop][0]) {
			columns.push({
				title: column,
				data: column,
				defaultContent: '<i style="color: red;">Data Missing</i>'
			});
		}

		$("#" + prop + "_table").DataTable({
			bSort: false,
			data: json[prop],
			columns: columns,
			responsive: true
		});
	}

};

if (window.File && window.FileList && window.FileReader && window.Blob) {

	$("#inputFile").on("change", function (e) {
		var inputFile = e.target.files[0];

		var fileReader = new FileReader();

		fileReader.onload = function (e) {
			var data = e.target.result;
			var workbook = XLSX.read(data, {
				type: 'binary'
			});

			var jsonData = toJSON(workbook);

			displayJSON(jsonData);

		};

		fileReader.readAsBinaryString(inputFile);

	});
} else {
	alert("HTML 5 File API not supportedd");
}

