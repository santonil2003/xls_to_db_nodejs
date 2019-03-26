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





pool.getConnection(function(err, connection) {

    try {

        var sql = `SELECT 
     concat('UPDATE text_history SET refund_id =',rtb.refund_text_id,' WHERE id=',th.id,' AND refund_id = 0') as update_query
FROM
    refund_text_bankfile rtb
        LEFT JOIN
    refund_text rt ON rtb.refund_text_id = rt.id
        LEFT JOIN 
    text_history th ON (th.id > 0
        AND RIGHT(rt.refund_mobile, 8) = RIGHT(th.mobile, 8))
WHERE
rt.amount_agreed = th.amout and th.refund_id = 0 and rt.status in ('3','4','5','6','8','9') 
ORDER BY rtb.refund_text_id`;



        // find the last entry 
        connection.query(sql, function(err, rows) {

            if (err) {
                console.log(err);
            }


            if (rows.length) {

                rows.forEach(function(row){

                    


                    connection.query(

                        row.update_query,

                        function(err, result) {

                            if (err) {
                                console.log(err);
                            }

                            // imported successfully...
                            console.log(row.update_query);
                            console.log('OK');
                        }
                    );


                });
            } else {
                console.log("Nothing to update..");
            }



        });





    } catch (err) {

        console.log(err);

    }

    //release the connection
    connection.release();
});


