/**
 * Copyright 2020 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Returns the index of the username that matches with the owner selected by the
 * user
 * @param {String} ownerSelection The username selected as owner by the user
 * @param {String[]} usernames String array containing usernames of all the
 *     engineers
 * @return {Number} Index of the matched username in the usernames array
 */
function match(ownerSelection, usernames) {
  for (var i = 0; i < usernames.length; i++) {
    if (usernames[i] == ownerSelection) {
      return i;
    }
  }
  return -1;
}

/**
 * Returns new date after addition of given number of days in the given date
 * @param {Date} date The date in which the given number of days are to be added
 * @param {Number} days The number of days to be added in the given date
 * @return {Date} New date after addition of given number of days in the given
 *     date
 */
function addDays(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
/**
 * Returns array of numbers containing coding units per day of each engineer
 * @param {Object[][]} engineerInfo A two dimensional array containing engineer
 *     information from the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {Number[]} Array of numbers containing coding units per day of each
 *     engineer
 */
function getCodingUnitsPerDay(engineerInfo, countEng) {
  var codingUnitsPerDay = [];
  for (var i = 0; i < countEng; i++) {
    if (engineerInfo[i][1] <= 0) {
      return;
    }
    var codingUnits = engineerInfo[i][1] / 7;
    codingUnitsPerDay.push(codingUnits);
  }
  return codingUnitsPerDay;
}

/**
 * Returns array of strings containing username of each engineer
 * @param {Object[][]} engineerInfo A two dimensional array containing engineer
 *     information from the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {String[]} Array of strings containing username of each engineer
 */
function getUsernames(engineerInfo, countEng) {
  var usernames = [];
  for (var i = 0; i < countEng; i++) {
    usernames.push(engineerInfo[i][0]);
  }
  return usernames;
}

/**
 * Returns empty array with given number of rows and columns
 * @param {Number} rowCount The number of rows required in the array
 * @param {Number} columnCount The number of columns required in the array
 * @return {Date[]} Empty array with given number of rows and columns
 */
function initializeArray(rowCount, columnCount) {
  var array = [];
  for (var i = 0; i < rowCount; i++) {
    var row = [];
    for (var j = 0; j < columnCount; j++) {
      row.push([]);
    }
    array.push(row);
  }
  return array;
}

/**
 * Returns an array containing initial task start dates of each engineer
 * Start dates displayed in the Tasks Sheet are initial task start dates
 * Initial Task Start Date for an engineer is the maximum of Project Start Date
 * and the Engineer's Start Date
 * @param {Date} projectStart Date object containing start date of the project
 * @param {Object[][]} engineerInfo A two dimensional array containing
 *     information of the engineers in the Team Sheet
 * @param {Number} countEng Number of engineers in the Team
 * @return {Date[]} Date array containing initial task start dates of each
 *     engineer
 */
function getDisplayStartDate(projectStart, engineerInfo, countEng) {
  var displayStartDateArray = [];
  for (var i = 0; i < countEng; i++) {
    var startDateEng = engineerInfo[i][2];
    if (projectStart < startDateEng) {
      displayStartDateArray.push(startDateEng);
    } else {
      displayStartDateArray.push(projectStart);
    }
  }
  return displayStartDateArray;
}

/**
 * Sets the task start dates (to be used for calculating the estimated
 * end date of the task for each engineer) in 'actualStartDateArray'.
 * The task start dates for these calculations is maximum of the
 * current date and the displayed task start date for each engineer (where
 * displayed task start date is maximum of project start date and engineer's
 * start date)
 * Sets corresponding value in 'isStartDateAfterTodayArray' array to
 * 'true' if the displayed task start date for the engineer is greater than
 * current date else 'false'
 * @param {Date[]} displayStartDateArray Date array containing displayed task
 *     start dates of each engineer
 * @param {Boolean[]} isStartDateAfterTodayArray Boolean array to store
 *     'true' if the displayed task start date is
 * greater than current date else 'false' for all engineers
 * @param {Date[]} actualStartDateArray Date array to store maximum of the
 *     current date and the displayed task start date for each engineer
 * @param {Number} countEng Number of engineers in the Team
 * @param {Date} currDate The current date
 */
function getActualStartDate(displayStartDateArray, isStartDateAfterTodayArray,
                            actualStartDateArray, countEng, currDate) {
  for (var i = 0; i < countEng; i++) {
    if (displayStartDateArray[i] > currDate) {
      actualStartDateArray.push(displayStartDateArray[i]);
      isStartDateAfterTodayArray.push(true);
    } else {
      actualStartDateArray.push(currDate);
      isStartDateAfterTodayArray.push(false);
    }
  }
}