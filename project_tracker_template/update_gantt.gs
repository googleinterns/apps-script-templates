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

const SERIAL_NUMBER_RANGE = 'Tasks!$A$7:$A';
const TITLE_RANGE = 'Tasks!$B$7:$B';
const START_DATE_RANGE = 'Tasks!$F$7:$F';
const EST_LAUNCH_DATE_RANGE = 'Tasks!$G$7:$G';
const STATUS_RANGE = 'Tasks!$H$7:$H';
const OWNER_RANGE = 'Tasks!$C$7:$C';
const TASK_NUMBER_RANGE = 'Tasks!$N$7:$N';
const USERNAME_RANGE = 'Team!$C$6:$C$25';
const OWNER_USERNAMES_RANGE = '$I$7:$I$26';
const COLOR_COLUMN_INDEX = 9;

/**
 * Updates all the custom formulas in the Timeline sheet
 * @param {Object[][]} ganttTableValues A two dimensional array containing
 *     information about tasks and milestones
 * @param {Number} numRows Number of rows containing milestone/ task data
 */
function updateGantt(ganttTableValues, numRows) {
  for (var i = 0; i < ganttTableValues.length; i++) {
    for (var j = 0; j < 13; j++) {
      ganttTableValues[i][j] = [ '' ];
    }
  }
  // Set Project Title
  ganttTableValues[0][1] = [ '=Tasks!B1' ];
  // Set Project Start Date
  ganttTableValues[0][2] = [ '=Tasks!H2' ];
  // Set Estimated Project Launch Date
  ganttTableValues[0][3] = [ '=Tasks!H4' ];
  // Set custom formula for week wise representation of the project timeline
  ganttTableValues[0][4] = [
    '=iferror(SPARKLINE({split(rept("7,",FLOOR((int(D6)-int(C6))/7)),","),mod(int(D6)-int(C6),7)},{"charttype","bar";"color1","white";"color2","lightblue"}),"")'
  ];
  ganttTableValues[0][8] = [ 'Milestone' ];
  // Set Task Serial Number from Task Sheet
  ganttTableValues[1][0] = [ '=ArrayFormula(' + SERIAL_NUMBER_RANGE + ')' ];
  // Set Task/ Milestone Title from Task Sheet
  ganttTableValues[1][1] = [ '=ArrayFormula(' + TITLE_RANGE + ')' ];
  // Set Task/ Milestone Start Date from Task Sheet
  ganttTableValues[1][2] = [ '=ArrayFormula(' + START_DATE_RANGE + ')' ];
  // Set Task/ Milestone Estimated Launch Date from Task Sheet
  ganttTableValues[1][3] = [ '=ArrayFormula(' + EST_LAUNCH_DATE_RANGE + ')' ];
  // Set Task/ Milestone Status from Task Sheet
  ganttTableValues[1][5] = [ '=ArrayFormula(' + STATUS_RANGE + ')' ];
  // Set Task/ Milestone Owner from Task Sheet
  ganttTableValues[1][6] = [ '=ArrayFormula(' + OWNER_RANGE + ')' ];
  // Set the usernames of team members in the Color Legend Block
  ganttTableValues[1][8] = [ '=ArrayFormula(' + USERNAME_RANGE + ')' ];
  // Set the Task Number in Column M (hidden) of the Tasks/ Milestones from the
  // Task Sheet
  ganttTableValues[1][12] = [ '=ArrayFormula(' + TASK_NUMBER_RANGE + ')' ];
  // Color Legend for Gantt Chart
  ganttTableValues[0][COLOR_COLUMN_INDEX] = [ '#9fc5e8' ];
  ganttTableValues[1][COLOR_COLUMN_INDEX] = [ '#3366CC' ];
  ganttTableValues[2][COLOR_COLUMN_INDEX] = [ '#DC3912' ];
  ganttTableValues[3][COLOR_COLUMN_INDEX] = [ '#FF9900' ];
  ganttTableValues[4][COLOR_COLUMN_INDEX] = [ '#43ff43' ];
  ganttTableValues[5][COLOR_COLUMN_INDEX] = [ '#990099' ];
  ganttTableValues[6][COLOR_COLUMN_INDEX] = [ '#ffc66f' ];
  ganttTableValues[7][COLOR_COLUMN_INDEX] = [ '#0099C6' ];
  ganttTableValues[8][COLOR_COLUMN_INDEX] = [ '#DD4477' ];
  ganttTableValues[9][COLOR_COLUMN_INDEX] = [ '#66AA00' ];
  ganttTableValues[10][COLOR_COLUMN_INDEX] = [ '#B82E2E' ];
  ganttTableValues[11][COLOR_COLUMN_INDEX] = [ '#316395' ];
  ganttTableValues[12][COLOR_COLUMN_INDEX] = [ '#994499' ];
  ganttTableValues[13][COLOR_COLUMN_INDEX] = [ '#22AA99' ];
  ganttTableValues[14][COLOR_COLUMN_INDEX] = [ '#AAAA11' ];
  ganttTableValues[15][COLOR_COLUMN_INDEX] = [ '#6633CC' ];
  ganttTableValues[16][COLOR_COLUMN_INDEX] = [ '#E67300' ];
  ganttTableValues[17][COLOR_COLUMN_INDEX] = [ '#8B0707' ];
  ganttTableValues[18][COLOR_COLUMN_INDEX] = [ '#329262' ];
  ganttTableValues[19][COLOR_COLUMN_INDEX] = [ '#5574A6' ];
  ganttTableValues[20][COLOR_COLUMN_INDEX] = [ '#3B3EAC' ];
  for (var i = 1; i <= numRows; i++) {
    var rowNumber = Number(i) + 6;
    // Evaluate color code for each task in Column L (hidden) using the color
    // legend and owner column
    ganttTableValues[i][11] =
        [ '=iferror(INDIRECT(if(iferror(MATCH($G' + rowNumber +
          ',ArrayFormula(' + OWNER_USERNAMES_RANGE +
          '),0),-1)>0,JOIN("","$J",TEXT(ADD(6,(MATCH($G' + rowNumber +
          ',ArrayFormula(' + OWNER_USERNAMES_RANGE +
          ')))),"###")))),"#9fc5e8")' ];
    // Create bar charts in Column E using the evaluated color codes in column L
    ganttTableValues[i][4] =
        [ '=iferror(SPARKLINE({int(C' + rowNumber + ')-int($C$6),int(D' +
          rowNumber + ')-int(C' + rowNumber +
          ')},{"charttype","bar";"color1","white";"color2",$L' + rowNumber +
          ';"max",int($D$6)-int($C$6)}),"")' ];
  }
}
