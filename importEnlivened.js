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
const sheetName = "Enlivened";



pool.getConnection(function (err, connection) {

    try {

        // find the last entry 
        connection.query("SELECT * FROM text_history ORDER BY id DESC LIMIT 1", function (err, rows) {

            if (err) {
                console.log(err);
            }


            if (rows.length) {

                var lastTextHistory = rows[0];

                // import into temp table
                importIntoTempTable(sheetName, lastTextHistory, connection);
            }



        });





    } catch (err) {

        console.log(err);

    }

    //release the connection
    connection.release();
});




// import sheet
function importIntoTempTable(sheetName, lastTextHistory, connection) {

    var filePath = __dirname + '/data/PCIM - AC00485.xlsx';
    var sheets = xlsx.parse(filePath);

    console.log("\n process sheet :" + sheetName);

    // parse sheet
    sheets.forEach(function (sheet) {

        if (sheetName == sheet.name.trim()) {

            console.log("\n Importing " + sheet.name);

            // rows from sheet
            var rows = sheet.data;

            // insert lock flag
            var insertLock = true;

            if (rows.length) {

                rows.forEach(function (row) {

                    // check if insert is locked..
                    if (!insertLock) {


                        // check if already imported
                        connection.query("SELECT * FROM mcom_enlivened WHERE payout_id ='" + row[0] + "'", function (err, rows) {
                            if (err) {
                                console.log(err);
                            }

                            // if not exist then good to insert..
                            if (rows.length == 0) {

                                var mcomEnlivened = {};
                                mcomEnlivened.payout_id = row[0];
                                mcomEnlivened.mobile = row[1];
                                mcomEnlivened.amount = row[2];
                                mcomEnlivened.enlivened = parseDateExcel(row[3]);

                                connection.query("INSERT INTO mcom_enlivened SET ?", mcomEnlivened, function (err, result) {

                                    if (err) {
                                        console.log(err);
                                    }

                                    if (result.insertId) {
                                        // import into text history...
                                        importIntoTextHistory(result.insertId, mcomEnlivened, connection);
                                    }

                                });



                            }
                        });









                    }


                    // realease insert lock once last payout_id is found in sheet
                    if (row.length && (row[0] == lastTextHistory.pay_id)) {
                        insertLock = false;
                    }

                });

            }

        }

    });


}




function importIntoTextHistory(mcomEnlivenedId, mcomEnlivened, connection) {

    var textHistory = {};
    textHistory.pay_id = mcomEnlivened.payout_id;
    textHistory.amout = mcomEnlivened.amount;
    textHistory.mobile = mcomEnlivened.mobile;
    textHistory.refundtype = "TEXT";
    textHistory.refund_id = "NULL";
    textHistory.enliven_date = mcomEnlivened.enlivened;

    connection.query("INSERT INTO text_history SET ?", textHistory, function (err, result) {

        if (err) {
            console.log(err);
        }

        if (result.insertId) {

            connection.query(
                'UPDATE mcom_enlivened SET imported=? WHERE id=?', ['1', mcomEnlivenedId],
                function (err, result) {
                    if (err) {
                        console.log(err);
                    }

                    // imported successfully...
                    console.log(textHistory);
                }
            );
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

/*
**

-- generate update query --
-- SELECT distinct(bankfilename) FROM refund_text_bankfile order by id desc limit 3;
-- k.id = highest text_history.id with text_history.refund_id > 0 from previous week

SELECT
    'UPDATE text_history SET refund_id ="',
    i.refund_text_id AS refund_id,
    '" WHERE id = ',
    k.id,
    ';',
    k.mobile, k.amout, j.amount_agreed, j.mobile
FROM
    refund_text_bankfile i
        LEFT JOIN
    refund_text j ON i.refund_text_id = j.id
        LEFT JOIN
    text_history k ON (k.id > 13346
        AND RIGHT(j.refund_mobile, 8) = RIGHT(k.mobile, 8))
WHERE
    i.bankfilename = '060717-1300refund-text.csv'
        AND j.amount_agreed = k.amout
ORDER BY k.id;
*/