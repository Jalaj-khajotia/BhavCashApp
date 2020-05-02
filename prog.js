console.log('Application loaded');

var XLSX = require('xlsx');
var fileName = '30.csv';
var minPrice = 10;
console.log('File used is ' + fileName);
var workbook = XLSX.readFile(fileName);
var sheet_name_list = workbook.SheetNames;

sheet_name_list.forEach(function (y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var excel = [];
    for (z in worksheet) {
        if (z[0] === '!')
            continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0, tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if (row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if (!excel[row])
            excel[row] = {};
        excel[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    excel.shift();
    excel.shift();

    mydata = [];
    var winners = [];
    var i = 0,
    limit = 0;

    const readline = require('readline').createInterface({
        input: process.stdin,
        output: process.stdout
    });

    readline.question(`Enter the 1. Daily Gainers 2. GapUp opener + price rise  \n \t  3. Bullish/Bearish Marbazu 4. Daily Loosers  \n \t  5. GapDown opener + Price looser`, (job) => {
        switch (job) {
        case '1':
            DailyGainers();
            break;
        case '2':
            GapUpGainers();
            break;
        case '3':
            Marbazu();
            break;
        case '4':
            DailyLoosers();
            break;
        case '3':
            GapDownLooser();
            break;

        default:
            console.log('kaboom');
        }
    });

    function DailyLoosers() {
        readline.question(`Enter the lower limit? `, (lower) => {
            readline.question('Enter Upper Limit, Default:-100% ', (upper) => {
                excel.forEach(function (cell) {
                    var fall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;

                    mydata[cell.SYMBOL] = Math.round(fall * 100) / 100;
                    limit = upper == 0 ? 100 : upper;
                    if (mydata[cell.SYMBOL] <= -lower && mydata[cell.SYMBOL] > -upper && cell.TOTTRDQTY > 10000 && cell.OPEN >= minPrice) {
                        // winners[cell.SYMBOL] = mydata[cell.SYMBOL];

                        winners[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: mydata[cell.SYMBOL],
                            CMP: cell.CLOSE
                        }
                    };
                });
                winners.sort(function (a, b) {
                    return b.Percentage - a.Percentage;
                });
                console.log('');
                console.log('Total no of stocks found are ' + winners.length);
                console.log('');
                console.log('Listing stocks which rose > ' + lower + '% but are lower than  < ' + upper + '%');
                console.log('');

                console.log('  ' + 'Stock Name' + '\t ' + '% Decrease' + '   CMP');
                winners.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t ' + stock.Percentage + '\t     ' + stock.CMP);
                })                
                readline.close()
            });
        });
    }

    function GapDownLooser() {
        gapUpList = [];
        var i = 0;
        readline.question(`Enter the Gap down %? `, (gapDown) => {
            readline.question('Enter Gainer minimum %, Default:0% ', (loss) => {
                excel.forEach(function (cell) {
                    var gapupPercentage = (cell.OPEN - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    var riseFall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    if (gapupPercentage < 0 && gapupPercentage > gapDown && riseFall >= loss && cell.TOTTRDQTY > 10000 && cell.OPEN >= minPrice) {

                        var percentage = Math.round(riseFall * 100) / 100;
                        var valuegapup = Math.round(gapupPercentage * 100) / 100;
                        gapUpList[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: percentage,
                            gapDown: valuegapup,
                            CMP: cell.CLOSE
                        };
                    }
                });
                console.log('');
                console.log('Total no of stocks found are ' + gapUpList.length);
                console.log('');
                console.log('Listing stocks with gap up > ' + gapDown + '% & gained > ' + loss + '%');
                console.log('');
                console.log('  ' + 'Stock Name' + '\t' + 'Gap Down %' + '\t' + 'Increase');
                gapUpList.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t' + stock.gapDown + '\t' + stock.Percentage);
                })
                readline.close();
            })
        })
    }

    function GapUpGainers() {
        gapUpList = [];
        var i = 0;
        readline.question(`Enter the Gap Up %? `, (gapUp) => {
            readline.question('Enter Gainer minimum %, Default:0% ', (gain) => {
                excel.forEach(function (cell) {
                    var gapupPercentage = (cell.OPEN - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    var riseFall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    if (gapupPercentage > 0 && gapupPercentage > gapUp && riseFall >= gain && cell.TOTTRDQTY > 10000 && cell.OPEN >= minPrice) {

                        var percentage = Math.round(riseFall * 100) / 100;
                        var valuegapup = Math.round(gapupPercentage * 100) / 100;
                        gapUpList[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: percentage,
                            GapUp: valuegapup,
                            CMP: cell.CLOSE
                        };
                    }
                });
                console.log('');
                console.log('Total no of stocks found are ' + gapUpList.length);
                console.log('');
                console.log('Listing stocks with gap up > ' + gapUp + '% & gained > ' + gain + '%');
                console.log('');
                console.log('  ' + 'Stock Name' + '\t' + 'GapUp %' + '\t' + 'Increase');
                gapUpList.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t' + stock.GapUp + '\t' + stock.Percentage);
                })
                readline.close();
            })
        })
    }

    function DailyGainers() {
        readline.question(`Enter the lower limit? `, (lower) => {
            readline.question('Enter Upper Limit, Default:100% ', (upper) => {
                excel.forEach(function (cell) {
                    var riseFall = (cell.CLOSE - cell.PREVCLOSE) * 100 / cell.PREVCLOSE;
                    mydata[cell.SYMBOL] = Math.round(riseFall * 100) / 100;
                    limit = upper == 0 ? 100 : upper;
                    if (mydata[cell.SYMBOL] >= lower && mydata[cell.SYMBOL] < upper && cell.TOTTRDQTY > 10000 && cell.OPEN >= minPrice) {
                        winners[i++] = {
                            Symbol: cell.SYMBOL,
                            Percentage: mydata[cell.SYMBOL],
                            CMP: cell.CLOSE
                        }
                    };
                });
                winners.sort(function (a, b) {
                    return a.Percentage - b.Percentage;
                });

                console.log('');
                console.log('Total no of stocks found are ' + winners.length);
                console.log('');
                console.log('Listing stocks which rose > ' + lower + '% but are lower than  < ' + upper + '%');
                console.log('');

                console.log('  ' + 'Stock Name' + '\t ' + '% Increase' + '   CMP');
                winners.forEach(function (stock) {
                    console.log('  ' + stock.Symbol + '\t ' + stock.Percentage + '\t     ' + stock.CMP);
                })

                readline.close()
            });
        });
    }

});