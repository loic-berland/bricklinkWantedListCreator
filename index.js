(async function () {

    const OAuth = require('oauth');
    const fs = require('fs');
    const XLSX = require("xlsx");

    const config = require('./creditentials.js');

    //OAuth to bricklink
    var oauth = new OAuth.OAuth('', '', config.consumer_key, config.consumer_secret, '1.0', null, 'HMAC-SHA1');

    //Opens the excel file
    var workbook = XLSX.readFile('test.xlsm');

    //Selects the first sheet and finds the length
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var arr = XLSX.utils.sheet_to_row_object_array(worksheet, { blankrows: false, defval: '' });
    const totalRows = arr.length + 1;

    //Starts the XML file
    fs.writeFileSync('wantedlist.xml', '<INVENTORY>\r\n');

    for (let i = 2; i < totalRows + 1; i++) {
        //bricklink API request to get a BL ref from a lego ID
        oauth.get('https://api.bricklink.com/api/store/v1/item_mapping/' + worksheet['A' + i].v, config.token, config.token_secret,
            function (error, data) {
                if (error) console.error(error);
                var list = JSON.parse(data);
                toXML(list, i, worksheet['A' + i].v);
            });
    }

    //Writes each part to an XML file
    function toXML(data, row, elementId) {
        if (Object.keys(data.data).length === 0) {
            console.log('error',elementId);
            fs.writeFileSync('badRefs.txt',elementId + "\r\n", {'flag' : 'a'});
        }
        else {
            fs.writeFileSync('wantedlist.xml', "   <ITEM>\r\n", { 'flag': 'a' });
            fs.writeFileSync('wantedlist.xml', "     <ITEMTYPE>P</ITEMTYPE>\r\n", { 'flag': 'a' });
            fs.writeFileSync('wantedlist.xml', '     <ITEMID>' + data.data[0].item.no + '</ITEMID>\r\n', { 'flag': 'a' });
            fs.writeFileSync('wantedlist.xml', '     <COLOR>' + data.data[0].color_id + '</COLOR>\r\n', { 'flag': 'a' });
            fs.writeFileSync('wantedlist.xml', '     <MINQTY>' + worksheet['B' + row].v + '</MINQTY>\r\n', { 'flag': 'a' });
            fs.writeFileSync('wantedlist.xml', "   </ITEM>\r\n", { 'flag': 'a' });
        }
    }
})()