/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values: number[][]): number {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
} 
 
 /**
    * Returns the number of distinct data values in the provided range
    * @customfunction
    * @param {any[][]} range The range to compute
    * @returns The count of unique values in this range
    */
   export function distinct(range): Promise<number> {
    if (typeof range[0][0] == typeof 'String' && range[0].length == 1) {
       var ctx = new Excel.RequestContext();
       var rangeObj = ctx.workbook.worksheets.getActiveWorksheet().getRange(range[0][0] as string);
       rangeObj.load("values");
       return (ctx.sync() as Promise<any>).then(() => {
         return getUniqueItems(rangeObj.values);
       })
    } else {
      return getUniqueItems(range as any[][]);
    }
  };
  function getUniqueItems(items: any[][]): Promise<number> {
    return new Promise(function (resolve, reject) {
      var uniqueVals = []
      try {
        for (var i = 0; i < items.length; i++) {
          for (var j = 0; j < items[i].length; j++) {
            if (uniqueVals.indexOf(items[i][j]) < 0) {
              uniqueVals.push(items[i][j]);
            }
          }
        }
      } catch (error) {
        console.error(error);
        reject(error);
      }
      resolve(uniqueVals.length);
    })
  }





/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
