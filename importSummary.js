const fs = require('fs');
const xlsx = require('node-xlsx');
const mysql = require('mysql');
const dateFormat = require('dateformat');
const config = require('./config.json');

const filePath = __dirname + '/summary_data/PCIM - AC00485.xlsx';



// custom params...
const pool = mysql.createPool({
	host: config.host,
	user: config.user,
	password: config.password,
	database: config.database
});



const sheets = ['Enlivened', 'Cashed', 'Cancelled', 'Expired'];

// start importing sheet data....
sheets.forEach(importALL);

function importALL(sheetName) {

	pool.getConnection(function (err, connection) {

		if (err) throw err;

		insertIntoDb(sheetName, connection);

		//release the connection
		connection.release();
	});
}



// import sheet
function insertIntoDb(sheetName, connection) {

	var sheets = xlsx.parse(filePath);

	console.log("\n process sheet :" + sheetName);

	// parse sheet
	sheets.forEach(function (sheet) {

		// read data for required sheet
		if (sheetName == sheet.name.trim()) {

			console.log("\n Importing " + sheet.name);

			// rows from sheet
			var rows = sheet.data;

			if (rows.length) {
				rows.forEach(function (row) {


					// check if already imported

					var mcomRefundSummary = {};
					mcomRefundSummary.sheet = sheetName;

					switch (sheetName) {
						case 'Enlivened':
							mcomRefundSummary.payout_id = row[0];
							mcomRefundSummary.mobile = row[1];
							mcomRefundSummary.amount = row[2];
							mcomRefundSummary.date = parseDateExcel(row[3]);
							break;
						case 'Cashed':
							mcomRefundSummary.payout_id = row[0];
							mcomRefundSummary.mobile = row[1];
							mcomRefundSummary.amount = row[2];
							mcomRefundSummary.date = parseDateExcel(row[4]);

							break;
						case 'Cancelled':
							mcomRefundSummary.payout_id = row[0];
							mcomRefundSummary.mobile = row[1];
							mcomRefundSummary.amount = row[2];
							mcomRefundSummary.date = parseDateExcel(row[4]);
							break;
						case 'Expired':
							mcomRefundSummary.payout_id = row[0];
							mcomRefundSummary.mobile = row[1];
							mcomRefundSummary.amount = row[2];
							mcomRefundSummary.date = parseDateExcel(row[4]);
							break;
					}


					connection.query("SELECT * FROM mcom_refund_summary WHERE payout_id=? AND sheet=?", [mcomRefundSummary.payout_id, mcomRefundSummary.sheet], function (err, rows) {
						if (err) throw err;

						// if not exist then good to insert..
						if (rows.length == 0) {

							if (mcomRefundSummary.payout_id && mcomRefundSummary.mobile && mcomRefundSummary.amount) {

								connection.query("INSERT INTO mcom_refund_summary SET ?", mcomRefundSummary, function (err, result) {
									if (err) throw err;
								});
							}


						}
					});


				});
			};





		}

	});


}


// excel date to mysql date
function parseDateExcel(serial) {

	try {

		var utc_days = Math.floor(serial - 25569);
		var utc_value = utc_days * 86400;
		var date_info = new Date(utc_value * 1000);

		var fractional_day = serial - Math.floor(serial) + 0.0000001;

		var total_seconds = Math.floor(86400 * fractional_day);

		var seconds = total_seconds % 60;

		total_seconds -= seconds;

		var hours = Math.floor(total_seconds / (60 * 60));
		var minutes = Math.floor(total_seconds / 60) % 60;

		var properdate = new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);

		return dateFormat(properdate, "yyyy-mm-dd");

	} catch (err) {
		return serial;
	}
}



