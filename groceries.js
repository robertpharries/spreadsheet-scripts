const _RECIPESHEET = 'Recipe';
const _RECIPENAMEROW = 2;
const _RECIPEOFFSET = 2;
const _RECIPEROWRANGE = 99;
const _RECIPEINGREDSTART = 4;
const _INGREDIENTMAX = 25;

const _WEEKLYSHEET = 'Weekly Selection List';
const _WEEKLYSTARTROW = 2;
const _WEEKLYSTARTCOL = 1;
const _WEEKLYHEADER = [
    ["Recipe Select", ""]
]

const _SHOPLISTSHEET = 'Shopping List';
const _SHOPLISTSTARTROW = 3;
const _SHOPLISTSTARTCOL = 1;
const _SHOPLISTHEADER = [
    ["Shopping List", "", "", ""],
    ["Quant.", "Units", "Item", "Cat."]
]

const _MEALPLANSHEET = 'Meal Plan';
const _MEALPLANSTARTROW = 1;
const _MEALPLANSTARTCOL = 1;
const _MEALPLANHEADER = [
    ["Shopping List", "", "", ""],
    ["Quant.", "Units", "Item", "Cat."]
]

const _SORTCONDITIONS = ["meat", "fruit", "veg", "condiment", "spice", "longlife", "fridge", "frozen"];

const _CONVERSIONS = {
    "cup": {
        "ml": 0.004,
        "tbsp": 0.08,
        "tsp": 0.02
    },
    "tbsp": {
        "tsp": 0.25
    },
    "kg": {
        "g": 0.001
    }
}

function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
        {name: 'Refresh Weekly Selection List', functionName: 'refreshWeekly'},
        {name: 'Refresh Ingredient Category List', functionName: 'refreshCategory'},
        {name: 'Generate Shopping List & Meal Plan', functionName: 'generateLists'}
    ];
    spreadsheet.addMenu('Directions', menuItems);
}

function onEdit() {
    var spreadsheet = SpreadsheetApp.getActive();
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);

    var numOfRecipe = getRecipeNameList(spreadsheet).length;
    recipeSheet.getRange(_RECIPEOFFSET, 3, _RECIPEROWRANGE, numOfRecipe*4)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setFontSize(10)
        .setFontWeight("normal");

    for (let i = 0; i < numOfRecipe; i++) {
        recipeSheet.setColumnWidth((i * 4) + 3, 52);
        recipeSheet.setColumnWidth((i * 4) + 4, 42);
        recipeSheet.setColumnWidth((i * 4) + 5, 160);
        recipeSheet.setColumnWidth((i * 4) + 6, 230);
    }
}
  
function refreshWeekly() {
    var spreadsheet = SpreadsheetApp.getActive();
    var weeklySheet = spreadsheet.getSheetByName(_WEEKLYSHEET);
    if (weeklySheet) {
        weeklySheet.clear();
    }
    else {
        weeklySheet = spreadsheet.insertSheet(_WEEKLYSHEET, spreadsheet.getNumSheets());
    }

    var recipeList = getRecipeNameList(spreadsheet);

    weeklySheet.getRange(_WEEKLYSTARTROW, _WEEKLYSTARTCOL, recipeList.length, 1).setValues(listToYRange(recipeList));

    //formatting
    weeklySheet.getRange(1, 1, 1, 2).setValues(_WEEKLYHEADER)
        .mergeAcross()
        .setFontSize(12)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setBackground("#cccccc")
        .setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
    weeklySheet.getRange(_WEEKLYSTARTROW, _WEEKLYSTARTCOL + 1, recipeList.length, 1).insertCheckboxes();
}

function generateLists() {
    var spreadsheet = SpreadsheetApp.getActive();
    var weeklySheet = spreadsheet.getSheetByName(_WEEKLYSHEET);

    if (!weeklySheet) {
        SpreadsheetApp.getUi().alert('No Weekly Sheet!');
        return;
    }

    generateShoppingList();
    generateMealPlan();
}

function generateShoppingList() {
    var spreadsheet = SpreadsheetApp.getActive();
    var shoplistSheet = spreadsheet.getSheetByName(generateShopListName());

    if (shoplistSheet) {
        shoplistSheet.clear();
        shoplistSheet.activate();
    }
    else {
        shoplistSheet = spreadsheet.insertSheet(generateShopListName(), spreadsheet.getNumSheets());
    }

    var recipesToBuy = getWeeklyCheckedRecipes(spreadsheet);
    var ingredientsToBuy = getIngFromRecipe(spreadsheet, recipesToBuy);
    if(ingredientsToBuy.length > 0) {
        var ingredientsToBuy = ingredientsToBuy.sort(function(a,b) {
            return (a[2] < b[2]) ? -1 : (a[2] > b[2]) ? 1 : 0;
        }).sort(function(a, b) {
            return (cond(a[3]) < cond(b[3])) ? -1 : (cond(a[3]) > cond(b[3])) ? 1 : 0;
        });
        shoplistSheet.getRange(_SHOPLISTSTARTROW, _SHOPLISTSTARTCOL, ingredientsToBuy.length, 4).setValues(ingredientsToBuy)
    }

    //formatting
    shoplistSheet.getRange(1, 1, 2, 4).setValues(_SHOPLISTHEADER);
    shoplistSheet.getRange(1, 1, 1, 4).mergeAcross()
        .setFontSize(12)
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    shoplistSheet.setColumnWidth(1, 52);
    shoplistSheet.setColumnWidth(2, 42);
    shoplistSheet.setColumnWidth(3, 360);
    shoplistSheet.setColumnWidth(4, 100);
    shoplistSheet.getRange(2, 1, 1, 4).setBorder(false, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function generateMealPlan() {
    var spreadsheet = SpreadsheetApp.getActive();
    var mealPlanSheet = spreadsheet.getSheetByName(generateMealPlanName());

    if (mealPlanSheet) {
        mealPlanSheet.clear();
    }
    else {
        mealPlanSheet = spreadsheet.insertSheet(generateMealPlanName(), spreadsheet.getNumSheets());
    }

    var recipesToBuy = getWeeklyCheckedRecipes(spreadsheet);

    var curX = _MEALPLANSTARTROW;
    for (let i = 0; i < recipesToBuy.length; i++) {
        var recipe = getRecipe(spreadsheet, recipesToBuy[i]);
        if(!recipe) {
            continue;
        }
        mealPlanSheet.getRange(curX, 1, recipe.length, recipe[0].length).setValues(recipe);
        
        var height = 0;
        for (let j = 0; j < recipe.length; j++) {
            if(!(recipe[j][0] || recipe[j][1] || recipe[j][2] || recipe[j][3])) {
                height = j;
                break;
            }
        }

        //formatting
        mealPlanSheet.getRange(curX + 1, 1, 1, 4).setBorder(true, false, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
        mealPlanSheet.getRange(curX + 1, 4, height-1, 1).setBorder(false, true, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
        mealPlanSheet.getRange(curX + 1, 4, 1, 1).setBorder(true, true, true, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
        mealPlanSheet.getRange(curX, 1, 1, 3).mergeAcross()
            .setFontSize(12)
            .setFontWeight("bold")
        mealPlanSheet.setColumnWidth(1, 52);
        mealPlanSheet.setColumnWidth(2, 42);
        mealPlanSheet.setColumnWidth(3, 360);
        mealPlanSheet.setColumnWidth(4, 1000);
        mealPlanSheet.getRange(curX, 4, 1, 1).setValues([["Day:"]])
            .setFontWeight("bold");
        mealPlanSheet.getRange("D1:D").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

        curX += (height + 1);
    }
}

function refreshCategory() {
    var spreadsheet = SpreadsheetApp.getActive();
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);

    var allIng = getIngFromRecipe(spreadsheet, getRecipeNameList(spreadsheet));
    var curIng = getIngCategory(spreadsheet);
    var allIngSep = [];

    for (let i = 0; i < allIng.length; i++) {
        if(curIng.indexOf(allIng[i][2]) < 0) {
            curIng.push(allIng[i][2]);
        }
        allIngSep.push(allIng[i][2]);
    }

    for (let i = 0; i < curIng.length; i++) {
        if(allIngSep.indexOf(curIng[i]) < 0) {
            curIng[i] = "- " + curIng[i];
        }
    }

    recipeSheet.getRange(_RECIPEINGREDSTART, 1, curIng.length, 1).setValues(listToYRange(curIng));

    curIng = recipeSheet.getRange(_RECIPEINGREDSTART, 1, curIng.length, 1).getValues();
    for (let i = 0; i < curIng.length; i++) {
        if(curIng[i][0].indexOf("- ") == 0) {
            curIng[i][0] = curIng[i][0].replace("- ", "");
            recipeSheet.getRange(_RECIPEINGREDSTART + i, 1, 1, 1).setFontLine("line-through");
        }
        else {
            recipeSheet.getRange(_RECIPEINGREDSTART + i, 1, 1, 1).setFontLine("none");
        }
    }
    recipeSheet.getRange(_RECIPEINGREDSTART, 1, curIng.length, 1).setValues(listToYRange(curIng));
}


// Util
function getRecipeNameList(spreadsheet) {
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);
    var result = [];

    if(recipeSheet) {
        var ranValues = recipeSheet.getRange(_RECIPENAMEROW, 1 + _RECIPEOFFSET, 1, _RECIPEROWRANGE * 4).getValues();
        for (let i = 0; i < ranValues[0].length; i+=4) {
            if(ranValues[0][i]) {
                result.push(ranValues[0][i]);
            }
        }
    }

    return result;
}

function getWeeklyCheckedRecipes(spreadsheet) {
    var weeklySheet = spreadsheet.getSheetByName(_WEEKLYSHEET);
    var result = [];

    var ranValues = weeklySheet.getRange(_WEEKLYSTARTROW, _WEEKLYSTARTCOL, _RECIPEROWRANGE, 2).getValues();

    for (let i = 0; i < ranValues.length; i++) {
        if(ranValues[i][1]) {
            result.push(ranValues[i][0]);
        }
    }

    return result;
}

function getRecipe(spreadsheet, recipeName) {
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);

    var orderedRecipeList = getRecipeNameList(spreadsheet);
    var index = orderedRecipeList.indexOf(recipeName);

    if(index >= 0) {
        return recipeSheet.getRange(2, (index * 4) + 1 + _RECIPEOFFSET, _INGREDIENTMAX, 4).getValues();
    }
    else {
        return null;
    }
}

function getIngFromRecipe(spreadsheet, recipeList) {
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);
    var result = [];

    var orderedRecipeList = getRecipeNameList(spreadsheet);

    var indexes = [];
    for (let i = 0; i < orderedRecipeList.length; i++) {
        if(recipeList.indexOf(orderedRecipeList[i]) >= 0) {
            indexes.push(i);
        }
    }

    for (let i = 0; i < indexes.length; i++) {
        var ranValues = recipeSheet.getRange(_RECIPEINGREDSTART, (indexes[i] * 4) + 1 + _RECIPEOFFSET, _INGREDIENTMAX, 3).getValues();

        for (let j = 0; j < ranValues.length; j++) {
            if(ranValues[j][0]) {
                result = pushIngredient(result, ranValues[j])
            }
        }
    }

    var ingCat = getIngWithCategory(spreadsheet);
    for (let i = 0; i < result.length; i++) {
        if(ingCat[result[i][2]]) {
            result[i].push(ingCat[result[i][2]]);
        }
        else {
            result[i].push("");
        }
    }

    return result;
}

function getIngCategory(spreadsheet) {
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);
    var result = [];

    var ranValues = recipeSheet.getRange(_RECIPEINGREDSTART, 1, 999, 2).getValues();
    for (let i = 0; i < ranValues.length; i++) {
        if(ranValues[i][0]) {
            result.push(ranValues[i][0]);
        }
        else {
            return result;
        }
    }

    return result;
}

function getIngWithCategory(spreadsheet) {
    var recipeSheet = spreadsheet.getSheetByName(_RECIPESHEET);
    var result = {};

    var ranValues = recipeSheet.getRange(_RECIPEINGREDSTART, 1, 999, 2).getValues();
    for (let i = 0; i < ranValues.length; i++) {
        if(ranValues[i][0]) {
            // result.push([ranValues[i][0], ranValues[i][1]]);
            result[ranValues[i][0]] = ranValues[i][1];
        }
        else {
            return result;
        }
    }

    return result;
}

function listToYRange(a_list) {
    var result = [];
    for (let i = 0; i < a_list.length; i++) {
        var intList = [];
        intList.push(a_list[i]);
        result.push(intList);
    }
    return result;
}

function generateShopListName() {
    return _SHOPLISTSHEET + " " + getDate();
}

function generateMealPlanName() {
    return _MEALPLANSHEET + " " + getDate();
}

function getDate() {
    var d = new Date();
    var month = '' + (d.getMonth() + 1);
    var day = '' + d.getDate();
    var year = d.getFullYear();

    if (month.length < 2) {
        month = '0' + month;
    }
    if (day.length < 2) {
        day = '0' + day;
    }

    return [day, month, year].join('-');
}

function pushIngredient(a_list, entry) {
    for (let i = 0; i < a_list.length; i++) {
        if(entry[2] == a_list[i][2] && entry[1] == a_list[i][1]) {
            a_list[i][0] += entry[0];
            return a_list;
        } else if(entry[2] == a_list[i][2]) {
            n = resolveUnits(a_list[i], entry);
            if(n) {
                a_list[i] = n;
                return a_list;
            }
        }
    }

    a_list.push(entry);
    return a_list;
}

function resolveUnits(old, entry) {
    var tryConversion = function(a, b) {
        if(_CONVERSIONS[a[1]]) {
            if(_CONVERSIONS[a[1]][b[1]]) {
                return _CONVERSIONS[a[1]][b[1]] * b[0];
            }
        }

        return null;
    }

    if(tryConversion(old, entry)) {
        var n = tryConversion(old, entry);
        old[0] = old[0] + n;
        return old;
    }
    else if(tryConversion(entry, old)) {
        var n = tryConversion(entry, old);
        old[0] = entry[0] + n;
        old[1] = entry[1];
        return old;
    }

    return null;
}

var cond = function(term) {
    var score;
    if (term == "") {
        score = _SORTCONDITIONS.length;
    } else {
        score = _SORTCONDITIONS.indexOf(term);
    }

    return score;
}
