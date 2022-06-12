(function main() {
	// global variable
	let strToLang = {
		"yyyy": "Années",
		"MM": "Mois",
		"dd": "Jours",
		"hh": "Heurs"
	}
	let obj = {
		idObj: parseInt({{[#CUROBJ]}}),
	}
	let btn = document.getElementById("EVA_TimeLine_btn");
	let popup = document.getElementById("EVA_TimeLine_popup");
	let table = popup.getElementsByTagName("table")[0];
	let popupInfo = document.getElementById("EVA_TimeLine_popupInfo");
	popupInfo.style.display = "none";
	
	let lsDiv = popup.querySelectorAll("#EVA_TimeLine_popup > div > div:nth-child(1) > div > div");
	let headerDiv = {
		creation: {
			title: lsDiv[0],
			when: lsDiv[1],
			who: lsDiv[2]
		},
		lastEdit: {
			title: lsDiv[3],
			when: lsDiv[4],
			who: lsDiv[5],
			what: lsDiv[6]
		},
		editBy: lsDiv[7]
	}
	
	let popupView = [];
	
	// functions
	let dateOutputFormat = (date1, date2) => {
		let isSameYear = date1.substring(0, 4) == date2.substring(0, 4);
		let isSameMonth = date1.substring(4, 6) == date2.substring(4, 6);
		let isSameDay = date1.substring(6, 8) == date2.substring(6, 8);
		
		if (isSameYear) {
			if (isSameMonth) {
				if (isSameDay) {
					return "hh";
				}
				return "dd";
			}
			return "MM";
		}
		return "yyyy";
	}
	let dateToReadable = (date, inputFormat, outputFormat) => {
		if (inputFormat !== "yyyyMMddhhmmss") return date;
		
		let splited = date.toString().split(/(..)/g);
		let year = splited[1] + splited[3];
		let month = splited[5];
		let day = splited[7];
		let hour = splited[9];
		let minute = splited[11];
		let second = splited[13];
		
		if (outputFormat == "dd/MM/yyyy hh:mm:ss") return `${day}/${month}/${year} ${hour}:${minute}:${second}`;
		if (outputFormat == "dd/MM/yyyy hh:mm") return `${day}/${month}/${year} ${hour}:${minute}`;
		if (outputFormat == "yyyy/MM/dd") return `${year}/${month}/${day}`;
		if (outputFormat == "yyyy/MM") return `${year}/${month}`;
		if (outputFormat == "hh") return `${hour}`;
		if (outputFormat == "dd") return `${day}`;
		if (outputFormat == "MM") return `${month}`;
		if (outputFormat == "yyyy") return `${year}`;
		if (outputFormat == "") return ``;
		return date;
	}
	let hidePopup = () => {
		popup.style.display = "none";
		popupInfo.style.display = "none";
	}
	let showPopup = () => {
		let json = {
			bAttr: true,
			reqType: "objData",
			idObj: obj.idObj,
			idForm: parseInt({{[#CURFORM]}}),
			bHisto: "true",
			session: EVA.cfg.idSess,
		};
		EVA.ut.wsCall(obj, json, showPopupCallBack);
	}
	let showPopupCallBack = () => {
		if (obj.jsonRes.errCode == 1) {
			// error web service
			alert("Le web service est inaccessible");
			return;
		}
		
		// fill edit list
		let lsEdit = [];
		obj.jsonRes.inputs.forEach(input => {
			if (input.values.length > 0) {
				input.values.forEach(value => {
					lsEdit.unshift({
						symbole: input.attr.FIELD,
						what: input.label,
						oldValue: "",
						newValue: value.val,
						who: value.whoCr.split("#")[1],
						when: value.dateCr,
						action: "Créer"
					});
					if (value.dateDel) {
						lsEdit.unshift({
							symbole: input.attr.FIELD,
							what: input.label,
							oldValue: value.val,
							newValue: "",
							who: value.whoDel.split("#")[1],
							when: value.dateDel,
							action: "Supprimer"
						});
					}
				});
			}
		});
		
		// sort lsEdit by date
		lsEdit.sort((a, b) => (a.when <= b.when) ? -1 : 1);
		
		// edit action to "modification" for each edit whith same when and what
		for (let i = 0; i < lsEdit.length - 1; i++) {
			if (lsEdit[i].when == lsEdit[i + 1].when && lsEdit[i].what == lsEdit[i + 1].what) {
				// remove i + 1 elem
				let rmEdit = lsEdit.splice(i + 1, 1);
				// replace i
				let newEdit = lsEdit[i];
				lsEdit[i].action = "Modification";
				lsEdit[i] = newEdit;
				if (lsEdit[i].oldValue == "") {
					if (rmEdit[0].oldValue == "") lsEdit[i].oldValue = rmEdit[0].newValue;
					else lsEdit[i].oldValue = rmEdit[0].oldValue;
				}
				else {
					if (rmEdit[0].oldValue == "") lsEdit[i].newValue = rmEdit[0].newValue;
					else lsEdit[i].newValue = rmEdit[0].oldValue;
				}
			}
		}
		
		popupView.push(lsEdit);
		
		updatePopup();
		popup.style.display = "flex";
	}
	let updatePopup = () => {
		let lsEdit = popupView[popupView.length - 1];
		// define who list
		let lsWho = lsEdit.map(edit => edit.who);
		// remove duplicated values in who list
		lsWho = [...new Set(lsWho)];
		
		// fill header
		headerDiv.lastEdit.who.innerText = lsEdit[lsEdit.length - 1].who;
		headerDiv.lastEdit.when.innerText = dateToReadable(lsEdit[lsEdit.length - 1].when, "yyyyMMddhhmmss", "dd/MM/yyyy hh:mm");
		headerDiv.lastEdit.what.innerText = "Champ: " + lsEdit[lsEdit.length - 1].what;
		
		headerDiv.creation.title.innerText = (popupView.length == 1) ? "Création de la fiche" : "Première modification";
		headerDiv.creation.who.innerText = lsEdit[0].who;
		headerDiv.creation.when.innerText = dateToReadable(lsEdit[0].when, "yyyyMMddhhmmss", "dd/MM/yyyy hh:mm");
		
		let lsWhat;
		if (lsWho.length == 1) {
			headerDiv.editBy.innerText = `Modifié par ${lsWho[0]}`;
			headerDiv.editBy.style.display = "block";
			
			// define what list
			lsWhat = lsEdit.map(edit => edit.symbole);
			// remove duplicated values in what list
			lsWhat = [...new Set(lsWhat)];
		}
		else {
			headerDiv.editBy.style.display = "none";
		}
		
		// fill table
		// group by date
		let dateFormat = dateOutputFormat(lsEdit[0].when.toString(), lsEdit[lsEdit.length - 1].when.toString());
		let lsColumn = {}
		lsEdit.map(edit => {
			edit.whenGrouped = dateToReadable(edit.when, "yyyyMMddhhmmss", dateFormat);
			if (lsColumn[edit.whenGrouped] !== undefined) lsColumn[edit.whenGrouped]++;
			else lsColumn[edit.whenGrouped] = 1;
		});
		
		let tableValue = [];
		(lsWhat ? lsWhat : lsWho).forEach(what => {
			let row = [what];
			Object.keys(lsColumn).forEach(key => {
				row.push(0);
			});
			tableValue.push(row);
		});
		let firstColDateFormat;
		if (dateFormat == "yyyy") firstColDateFormat = "";
		else firstColDateFormat = "yyyy/MM/dd".split("/" + dateFormat)[0];
		let lastRow = [`${dateToReadable(lsEdit[0].when, "yyyyMMddhhmmss", firstColDateFormat)} ${strToLang[dateFormat]}:`];
		// fill last row
		let columnKey = Object.keys(lsColumn).sort((a, b) => (a * 1 < b * 1) ? -1 : 1);
		columnKey.forEach(key => {
			lastRow.push(key);
		});
		tableValue.push(lastRow);
		// insert edit count
		lsEdit.forEach(edit => {
			let y = lsWhat ? lsWhat.indexOf(edit.symbole) : lsWho.indexOf(edit.who);
			let x = columnKey.indexOf(edit.whenGrouped);
			tableValue[y][x + 1]++;
		});
		
		// show table value in table
		// empty table
		table.replaceChildren();
		tableValue.forEach((row, i) => {
			let tr = document.createElement("tr");
			row.forEach((value, j) => {
				let td = document.createElement("td");
				
				if (j > 0) {
					let div = document.createElement("div");
					div.style.display = "flex";
					let span = document.createElement("span");
					span.style.margin = "auto";
					span.innerText = value;
					if (i !== tableValue.length - 1) {
						// not last row
						span.style.color = "white";
						if (span.innerText !== "0") td.style.background = "green";
					}
					div.appendChild(span);
					td.appendChild(div);
				}
				else {
					// first col
					td.innerText = value;
					if (i == tableValue.length - 1) td.style.fontWeight = "bold";
				}
				
				// set click event on cells
				if (td.innerText !== "" && td.innerText !== "0") {
					let newLsEdit;
					let timeoutId;
					if (i !== tableValue.length - 1 || j !== 0) {
						// generate new edit list
						let rowName = tableValue[i][0];
						let colName = tableValue[tableValue.length - 1][j];
						if (j == 0) {
							newLsEdit = lsEdit.filter(edit => (lsWhat ? edit.symbole : edit.who) == rowName);
						}
						else if (i == tableValue.length - 1) {
							newLsEdit = lsEdit.filter(edit => edit.whenGrouped == colName);
						}
						else {
							newLsEdit = lsEdit.filter(edit => (lsWhat ? edit.symbole : edit.who) == rowName && edit.whenGrouped == colName);
						}
						
						// set title
						td.addEventListener("mouseenter", evt => {
							clearTimeout(timeoutId);
							timeoutId = setTimeout(() => {
								// show info
								
								// fill table
								let tbody = popupInfo.getElementsByTagName("tbody")[0];
								// empty tbody
								tbody.replaceChildren();
								newLsEdit.forEach(edit => {
									// add tr in tbody
									let tr = document.createElement("tr");
									
									// define value in td
									let lsAttrValue = ["oldValue", "newValue"];
									lsAttrValue.forEach(attr => {
										let newAttr = attr + "Td";
										edit[newAttr] = edit[attr].replaceAll("\n", " ");
										if (edit[newAttr].length > 15) {
											edit[newAttr] = edit[newAttr].substring(0, 10) + ` (${edit[newAttr].length})`;
										}
									});
									
									// create td
									let lsTdValue = [
										edit.who,
										edit.action,
										edit.what,
										edit.oldValueTd,
										edit.newValueTd,
										dateToReadable(edit.when, "yyyyMMddhhmmss", "dd/MM/yyyy hh:mm:ss"),
									];
									lsTdValue.forEach(tdValue => {
										let td = document.createElement("td");
										td.innerText = tdValue;
										tr.appendChild(td);
									});
									
									tbody.appendChild(tr);
								});
								
								// position
								popupInfo.style.top = popup.children[0].getBoundingClientRect().bottom + "px";
								
								popupInfo.style.display = "flex";
							}, 1000);
						});
						td.addEventListener("mouseleave", evt => {
							clearTimeout(timeoutId);
						});
					}
					td.addEventListener("click", () => {
						clearTimeout(timeoutId);
						
						if (newLsEdit) {
							// zoom in
							if (newLsEdit.length !== lsEdit.length) popupView.push(newLsEdit); ////////////////////////////////////////////////
						}
						else {
							// zoom out
							if (popupView.length > 1) popupView.pop();
						}
						
						updatePopup();
					});
					td.style.cursor = "pointer";
				}
				
				tr.appendChild(td);
			});
			table.appendChild(tr);
		});
	}
	
	// set popup
	popup.remove();
	hidePopup();
	document.body.appendChild(popup);
	
	// set events
	btn.addEventListener("click", () => {
		showPopup();
	});
	popup.addEventListener("click", () => {
		hidePopup();
	});
	popup.getElementsByTagName("i")[0].addEventListener("click", evt => {
		hidePopup();
	});
	popup.getElementsByTagName("div")[0].addEventListener("click", evt => {
		evt.stopPropagation();
		popupInfo.style.display = "none";
	});
	popupInfo.addEventListener("click", evt => {
		evt.stopPropagation();
	});
})();