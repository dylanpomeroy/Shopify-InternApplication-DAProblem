var xlsx = require('node-xlsx')

// open excel file
var excelObj = xlsx.parse('./data.xlsx')

// get data on Sheet1 (first)
var data = excelObj[0].data;
data.shift()

// variables used in execution
var quantityCounts = {}
var quantityWeights = {}
var totalCount = data.length - 1
var totalSpentWeighted = 0
var averageSpentWeighted = 0

// gather purchased amount counts for weighting calculations
data.forEach(function(element) {
    if (element[4] in quantityCounts) quantityCounts[element[4]] += 1
    else quantityCounts[element[4]] = 1;
}, this)

// calculate weights for each count
for (var count in quantityCounts){
    quantityWeights[count] = quantityCounts[count] / totalCount
}

// find total spent after applying weights to each order totals
data.forEach(function(element){
    totalSpentWeighted +=  quantityWeights[element[4]] * element[3]
}, this)

// find average of weighted spendings
averageSpentWeighted = totalSpentWeighted / totalCount

console.log("Average Weighted Order Value: "+averageSpentWeighted)