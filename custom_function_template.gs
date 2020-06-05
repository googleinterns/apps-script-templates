/**
* A custom function template for the following functions - 
*
* (operation = 0) Get sum of given values 
* @return The sum of given values 
*
* (operation = 1) Get mean of given values 
* @return The mean of given values
*
* (operation = 2) Get variance for the given input values.
* @return The variance for the given input values.
*
* (operation = 3) Get standard deviation for the given input values.
* @return The standard deviation for the given input values.
*
* (operation = 4) Get all the modes for the given input values.
* @return The modes for the given input values.
*
* (operation = 5) Get maximum value for the given values.
* @return The maximum value for the given values.
*
* (operation = 6) Get countIf - count of input values equal to the given value 
* @return The count of values equal to the first input value.
*
* (operation = 7) Get The median for the given input values.
* @return The median for the given input values.
*
* @customfunction
*/
function customFunctionTemplate() {
  if (arguments.length < 2) {
    throw new Error("Too few arguments");
  }
  var operation = arguments[0]; // First argument determines the operation to be performed.
  var values = Array.prototype.slice.call(arguments).slice(1); // Input values are second argument onwards. 
  var len = values.length; // Length of input values (on which function is to be performed).
  
  
  // Calculates sum of input values.
  if (operation === 0) {
    var sum = 0;
    for (var i = 0 ; i < len ; i++) {
      sum += values[i];
    }
    return sum;
  }
  
  // Calculates mean of input values 
  if (operation === 1) {
    var sum = 0;
    for (var i = 0 ; i < len ; i++) {
      sum += values[i];
    }
    var mean = sum/len;
    return mean;
  }
  
  // Calculates variance of input values 
  if (operation === 2) {
    var sum = 0;
    for (var i = 0 ; i < len ; i++) {
      sum += values[i];
    }
    var mean = sum/len;
    var sumSquareDifference = 0; // Sum of squared differences
    // Summation of squared difference between each input value and mean of all input values.
    for (var i = 0; i < len; i++) {
      sumSquareDifference+= Math.pow((values[i] - mean),2);
    }
    var variance = sumSquareDifference/len; // Variance is mean of sum of squared differences.
    return variance;
  }
  
  // Calculates standard deviation of input values
  if (operation === 3) {
    var sum = 0;
    for (var i = 0 ; i < len ; i++) {
      sum += values[i];
    }
    var mean = sum/len;
    var sumSquareDifference = 0; 
    for (var i = 0; i < len; i++) {
      sumSquareDifference+= Math.pow((values[i] - mean),2);
    }
    var variance = sumSquareDifference/len; 
    // Standard deviation is square root of variance.
    var standardDeviation = Math.sqrt(variance); 
    return standardDeviation;
  }
  
  // Calculates mode of input values
  if (operation === 4) {
    var valueFrequency = {}; // Mapping frequency of each input value
    var maxCount = 0; // Maximum frequency observed
    var modes = []; // All input values that have frequency = maxCount
    values.forEach( function(val) {
      // If frequency of current value is 0, increment to 1.  
      if (!valueFrequency[val]) {
        valueFrequency[val] = 1;
      } else {
        valueFrequency[val]++; // Increment frequency.
      }
      // If frequency of current value is greater than maximum frequency observed,
      // current value is the new mode and update maximum frequency. 
      if (valueFrequency[val] > maxCount) {
        modes = [val];
        maxCount = valueFrequency[val];
	  } else if (valueFrequency[val] === maxCount) {
        modes.push(val); // A value is a mode if its frequency is equal to maximum observed frequency.
        maxCount = valueFrequency[val];
	  }
    });
    return modes; 
  }
  
  // Calculates maximum value for the given input values.
  if (operation === 5) {
    var maxValue = Number.MIN_VALUE;
    values.forEach( function(val) {
      maxValue = Math.max(maxValue, val);
    });
    return maxValue; 
  }

  // Calculates countIf - count of input values equal to the given value 
  if (operation === 6) {
    var countIf = 0;
    var conditionValue = values[0]; // Value for equality condition.
    var inputValues = values.slice(1); // Input values to compare with conditionValue.
    var inputValuesLength = len-1;
    for (var i = 0; i < inputValuesLength; i++) {
      if (conditionValue === inputValues[i]) {
        countIf++;
      }
    }
    return countIf;
  }
  
  // Calculates median of input values  
  if (operation === 7) {
    var median = 0;
    var remainder = len % 2;
    var mid = (len - remainder) / 2; // Find middle index
    values.sort(); // Sort the array
    if(remainder === 1) {
      // Median is middle element in array of odd length
      median = values[mid];  
    } else {
      // Median is average of the 2 middle elements in array of even length
      median = (values[mid - 1] + values[mid]) / 2; 
    }
    return median;
  }
}