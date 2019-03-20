
Step 1: Download the PCIM - AC00485.xlsx from the email

Step 2: Copy the downloaded file into the mcom/data (at a moment full path to data is /var/www/html/mcom/data)

Step 3: Run importEnlivened.js 

this script will import the rows from enlivened tab of the  xlsx sheet into the text_history table.

Now, we need to find the refund_id from refund_text, for imported data. 

Follow the given steps to update the refund_id in text_history table.


Step 4 : find the list of bankfile using

SELECT distinct(bankfilename) FROM refund_text_bankfile order by id desc limit 10;


Step 5 : find the last id of the text_history aka k.id


Step 6:  Prepare the following query based on Step 4 and Step 5

SELECT 
    'UPDATE text_history SET refund_id ="',
    i.refund_text_id AS refund_id,
    '" WHERE id = ',
    concat(k.id,';')
FROM
    refund_text_bankfile i
        LEFT JOIN
    refund_text j ON i.refund_text_id = j.id
        LEFT JOIN 
    text_history k ON (k.id > 16991
        AND RIGHT(j.refund_mobile, 8) = RIGHT(k.mobile, 8))
WHERE i.bankfilename IN ('210219-1300refund-text.csv',
'140219-1300refund-text.csv',
'070219-1300refund-text.csv',
'310119-1300refund-text.csv',
'240119-1300refund-text.csv',
'170119-1300refund-text.csv',
'100119-1300refund-text.csv',
'020119-1300refund-text.csv') AND 
j.amount_agreed = k.amout and k.refund_id = 0
ORDER BY i.refund_text_id;

Note: k.id is obtained from Step 5.
And, i.bankfilename is obtained from step 4. Make sure you use only the unsed bankfile name from the step 4 or you may overlap the last one from the previous import.


Step 6 : make sure update query prepared from step 6 does not update same text_history row... you will need to check the list of update query manually to avoid it.

Step 7: once you update the refund_id in text_history run following commands

$ node importCancelled.js
$ node importCashed.js
$ node importExpired.js



Repeat these steps for every new PCIM - AC00485.xlsx from mcom









