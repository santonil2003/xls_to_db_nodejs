var fs = require('fs');
var xlsx = require('node-xlsx');
var mysql = require('mysql');
var dateFormat = require('dateformat');
var config = require('./config.json');



// custom params...
var pool = mysql.createPool({
    host: config.host,
    user: config.user,
    password: config.password,
    database: config.database
});



// sheet constants
const sheetName = "Cashed";



pool.getConnection(function(err, connection) {

    try {

        // find the last entry 
        connection.query("SELECT * FROM text_history_status WHERE statustype='CASH' ORDER BY id DESC LIMIT 1", function(err, rows) {

            if (err) {
                console.log(err);
            }


            if (rows.length) {

                var lastTextHistoryStatus = rows[0];

                // import into temp table
                importIntoTempTable(sheetName, lastTextHistoryStatus, connection);
            }



        });





    } catch (err) {

        console.log(err);

    }

    //release the connection
    connection.release();
});




// import sheet
function importIntoTempTable(sheetName, lastTextHistoryStatus, connection) {

    var filePath = __dirname + '/data/PCIM - AC00485.xlsx';
    var sheets = xlsx.parse(filePath);

    console.log("\n process sheet :" + sheetName);

    // parse sheet
    sheets.forEach(function(sheet) {

        if (sheetName == sheet.name.trim()) {

            console.log("\n Importing " + sheet.name);

            // rows from sheet
            var rows = sheet.data;

            // insert lock flag
            var insertLock = true;

            if (rows.length) {

                rows.forEach(function(row) {


                    if (row.length) {

                        // check if insert is locked..
                        if (!insertLock) {


                            // check if already imported
                            connection.query("SELECT * FROM mcom_cahsed WHERE payout_id ='" + row[0] + "'", function(err, rows) {
                                if (err) {
                                    console.log(err);
                                }

                                // if not exist then good to insert..
                                if (rows.length == 0) {

                                    var mcomCashed = {};
                                    mcomCashed.payout_id = row[0];
                                    mcomCashed.mobile = row[1];
                                    mcomCashed.amount = row[2];
                                    mcomCashed.cashed = row[3];
                                    mcomCashed.cashed_date = parseDateExcel(row[4]);

                                    connection.query("INSERT INTO mcom_cahsed SET ?", mcomCashed, function(err, result) {

                                        if (err) {
                                            console.log(err);
                                        }

                                        if (result.insertId) {
                                            // import into text history...
                                            importIntoTextHistory(result.insertId, mcomCashed, connection);
                                        }

                                    });



                                }
                            });
                        }


                        // realease insert lock once last payout_id is found in sheet
                        if (row.length && (row[0] == lastTextHistoryStatus.pay_id)) {
                            insertLock = false;
                        }

                    }

                });

            }

        }

    });


}




function importIntoTextHistory(mcomCashedId, mcomCashed, connection) {

    var textHistoryStatus = {};
    textHistoryStatus.pay_id = mcomCashed.payout_id;
    textHistoryStatus.amount = mcomCashed.amount;
    textHistoryStatus.mobile = mcomCashed.mobile;
    textHistoryStatus.statustype = "CASH";
    textHistoryStatus.change_date = mcomCashed.cashed_date;
    textHistoryStatus.is_booked = "0";

    connection.query("INSERT INTO text_history_status SET ?", textHistoryStatus, function(err, result) {

        if (err) {
            console.log(err);
        }

        if (result.insertId) {


            // find refund_id
            connection.query("SELECT * FROM text_history WHERE pay_id ='" + textHistoryStatus.pay_id + "' ORDER BY id DESC LIMIT 1 ", function(err, rows) {

                if (err) {
                    console.log(err);
                }

                if (rows.length > 0) {

                    var textHistory = rows[0];


                    // update refund_id
                    connection.query(
                        'UPDATE refund_text set status=? WHERE id=?', ['4', textHistory.refund_id],
                        function(err, result) {
                            if (err) {
                                console.log(err);
                            }




                            // marke as updated
                            connection.query(
                                'UPDATE mcom_cahsed SET imported=? WHERE id=?', ['1', mcomCashedId],
                                function(err, result) {
                                    if (err) {
                                        console.log(err);
                                    }

                                    // imported successfully...
                                    textHistoryStatus.refund_id = textHistory.refund_id;
                                    console.log(textHistoryStatus);
                                }
                            );







                        }
                    );



                }
            });





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



// debug
function r(val) {
    console.log("\n");
    console.log(val);
}