var cheerio = require('cheerio');
var request = require('request');
var XLSX = require('xlsx');
var wait = require('wait.for');

var infile:string = process.argv[2];
var outfile: string = process.argv[3];
var useSheetName: string = process.argv[4];
var init: string = process.argv[5];
var workbook = XLSX.readFile(infile, { cellDates: true });

//here we use wait.for. The launchfiber calls the function with which we can serialize with a wait.for. While we can parallelize and get multiple quotes
// at the same time, this takes the approach of getting a quote at a time
wait.launchFiber(updateFileWithQuotes);

function updateFileWithQuotes()
{
    processQuotes();
    //we are all done with all the quotes
    // now write the file out.
    XLSX.writeFile(workbook, outfile, { cellDates: true, bookType: "xlsx" });
}

function processQuotes() {
    
    var hdrSym:string = 'Symbol';
    var hdrWeb:string = 'WebPrice';
    var symbol:string;
    var price:string;
    
    var workSheet = workbook.Sheets[useSheetName];
    
    var colAlpha=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];
    var colHdrs={};
    //Row 1.
    
    
    //From A to M columns, populate the colHdrs array with the name of the column
    for (var cols=0; cols<colAlpha.length; cols++) {
        if (workSheet[colAlpha[cols]+1] !== undefined) {
            colHdrs[workSheet[colAlpha[cols]+1].v] = colAlpha[cols];
        }
    }
    
    //now for every other row, find the symbol column and webprice column
    

    for (var rows = 2; rows < 10000; rows++) {
        //continue till we find a row without a symbol
        if (workSheet[colHdrs[hdrSym] + rows] === undefined) {
            break;
        }
        //get the symbol
        symbol = workSheet[colHdrs[hdrSym] + rows].v;

        //if we are init, we want all quotes, so set value to 0 so we know when dont get a quote later on

        if (init == "init") {
            workSheet[colHdrs[hdrWeb] + rows] = { t: 'n', v: 0 };
        }
        // init is not set to init, so we are in the repeat run, only get quotes for ones we dont already have a quote
        else if (workSheet[colHdrs[hdrWeb] + rows].v != 0)
        {
            continue;
        }

        // now, call the async function "get_a_quote" using wait.for, so it becomes sync.
        // if the async returns error, wait.for will throw.
        try {
            price = wait.for(get_a_quote, symbol);
            //there are some quotes which will be with a , (like google give 1,122)
            //strip the commas in the quote and then use parseFloat to convert the string to a number
            var joinPrice: string = price.split(',').join('');

            var newPrice: number = parseFloat(joinPrice);

            console.log("price is " + newPrice);
            //store the price in the WebPrice column
            workSheet[colHdrs[hdrWeb] + rows] = { t: 'n', v: newPrice };
        //console.log("price for " + workSheet[colHdrs[hdrSym]+rows].v + "is " + workSheet[colHdrs[hdrWeb]+ rows].v);   
        } catch (err) {
            console.log("got an error for symbol " + symbol + " error is " + err);
            price = '0';
        }
    }
}


function get_a_quote(ticker:String,  callback) {

    var url:string = 'http://money.cnn.com/quote/quote.html';
    
    //console.log("going to request for symbol " + ticker);

    const options = {
      method: 'GET',
      uri: url,
      qs: {
        symb:ticker
      }
    }
    request(options, function (err, response, body) {
    
        if (err) {
            console.log("Got error " + err);
        }
    
    
        // Tell Cherrio to load the HTML
        $ = cheerio.load(body);
        
        
            //cnn returns the price within a class called .wsod_quoteData. That element has a starting span within which there is the price, and it has other
        //things. So we find the first table data, and then find first span and get the text within it.
        var price = $('.wsod_quoteData').find('td').first().find('span').first().text();
        
        if (price) { 
            // console.log("price for " + ticker + "is " + $('.wsod_quoteData').html());

        }
        else {
            // If we didnt find the price, it is elsewhere. In case of mutual funds, there is no span. It is right in the table between table data and 
            // a div. So we find that segment. Then it gets tricky, because the element has price as well as additional children html. So we go back
            // one element, clone it, remove the children, and get the text to get the price.
            price = $('.wsod_quoteData').find('td').first().nextUntil('div').prev().clone().children().remove().end().text();
            // console.log("price for " + ticker + "is " + $('.wsod_quoteData').find('td').first().nextUntil('div').prev().clone().children().remove().end().text());
        }
        
        if (price) {
             console.log("price for " + ticker + " is " + price);
            callback(0, price);
        }
        else {
            console.log("still no price for " + ticker);
            callback("no price");
            //console.log("body is " + body);
        }
        
    });
}
