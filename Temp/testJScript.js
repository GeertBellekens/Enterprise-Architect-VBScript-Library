//[group=Temp]
!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: 
 * Author: 
 * Purpose: 
 * Date: 
 */
 

   


/**
 * String.prototype.at()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support    No      No      No                  No    No      No
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.at) {
  Object.defineProperty(String.prototype, "at",
    {
      value: function (n) {
        // ToInteger() abstract op
        n = Math.trunc(n) || 0;
        // Allow negative indexing from the end
        if (n < 0) n += this.length;
        // OOB access is guaranteed to return undefined
        if (n < 0 || n >= this.length) return undefined;
        // Otherwise, this is just normal property access
        return this[n];
      },
      writable: true,
      enumerable: false,
      configurable: true
    });
}

/**
 * String.fromCharCode()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.fromCodePoint()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41  	29      (No)	            28	    10      ?
 * -------------------------------------------------------------------------------
 */


/**
 * String.anchor()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	1.0     (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.charAt()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.charCodeAt()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	1.0     (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.codePointAt()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41  	29      11	                28	    10      ?
 * -------------------------------------------------------------------------------
 */


/**
 * String.concat()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.endsWith()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41  	17      (No)	            (No)	9       (Yes)
 * -------------------------------------------------------------------------------
 */


/**
 * String.includes()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41  	40      (No)	            (No)	9       (Yes)
 * -------------------------------------------------------------------------------
 */


/**
 * String.indexOf()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.lastIndexOf()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.link()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	1.0    (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.localeCompare()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	1.0    (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.match()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.normalize()
 * version 0.0.1
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	34   	31      (No)	            (Yes)	10      (Yes)
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.normalize) {
  // need polyfill
}

/**
 * String.padEnd()
 * version 1.0.1
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	57   	48      (No)	            44   	10      15
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.padEnd) {
  Object.defineProperty(String.prototype, 'padEnd', {
    configurable: true,
    writable: true,
    value: function (targetLength, padString) {
      targetLength = targetLength >> 0; //floor if number or convert non-number to 0;
      padString = String(typeof padString !== 'undefined' ? padString : ' ');
      if (this.length > targetLength) {
        return String(this);
      } else {
        targetLength = targetLength - this.length;
        if (targetLength > padString.length) {
          padString += padString.repeat(targetLength / padString.length); //append to original to ensure we are longer than needed
        }
        return String(this) + padString.slice(0, targetLength);
      }
    },
  });
}

/**
 * String.padStart()
 * version 1.0.1
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	57   	51      (No)	            44   	10      15
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.padStart) {
  Object.defineProperty(String.prototype, 'padStart', {
    configurable: true,
    writable: true,
    value: function (targetLength, padString) {
      targetLength = targetLength >> 0; //floor if number or convert non-number to 0;
      padString = String(typeof padString !== 'undefined' ? padString : ' ');
      if (this.length > targetLength) {
        return String(this);
      } else {
        targetLength = targetLength - this.length;
        if (targetLength > padString.length) {
          padString += padString.repeat(targetLength / padString.length); //append to original to ensure we are longer than needed
        }
        return padString.slice(0, targetLength) + String(this);
      }
    },
  });
}

/**
 * String.repeat()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41   	24      (No)	            (Yes)   9       (Yes)
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.repeat) {
  Object.defineProperty(String.prototype, 'repeat', {
    configurable: true,
    writable: true,
    value: function (count) {
      if (this == null) {
        throw new TypeError("can't convert " + this + ' to object');
      }
      var str = '' + this;
      count = +count;
      if (count != count) {
        count = 0;
      }
      if (count < 0) {
        throw new RangeError('repeat count must be non-negative');
      }
      if (count == Infinity) {
        throw new RangeError('repeat count must be less than infinity');
      }
      count = Math.floor(count);
      if (str.length == 0 || count == 0) {
        return '';
      }
      if (str.length * count >= 1 << 28) {
        throw new RangeError(
          'repeat count must not overflow maximum string size'
        );
      }
      var rpt = '';
      for (; ;) {
        if ((count & 1) == 1) {
          rpt += str;
        }
        count >>>= 1;
        if (count == 0) {
          break;
        }
        str += str;
      }
      return rpt;
    },
  });
}

/**
 * String.search()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.slice()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.split()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.startsWith()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41   	17      (No)	            28   	9       (Yes)
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.startsWith) {
  Object.defineProperty(String.prototype, 'startsWith', {
    configurable: true,
    writable: true,
    value: function (searchString, position) {
      position = position || 0;
      return this.substr(position, searchString.length) === searchString;
    },
  });
}

/**
 * String.substr()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.substring()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.toLocaleLowerCase()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.toLocaleUpperCase()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.toLowerCase()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.toString()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.toUpperCase()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)	            (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.trim()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	3.5     9    	            10.5	5       ?
 * -------------------------------------------------------------------------------
 */
if (!String.prototype.trim) {
  Object.defineProperty(String.prototype, 'trim', {
    configurable: true,
    writable: true,
    value: function () {
      return this.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
    },
  });
}

/**
 * String.trimLeft()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	3.5     (No)    	        ?	    ?       ?
 * -------------------------------------------------------------------------------
 */

/**
 * String.trimRight()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	3.5     (No)    	        ?	    ?       ?
 * -------------------------------------------------------------------------------
 */

/**
 * String.valueOf()
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	(Yes)  	(Yes)   (Yes)    	        (Yes)	(Yes)   (Yes)
 * -------------------------------------------------------------------------------
 */

/**
 * String.raw
 * version 0.0.0
 * Feature	        Chrome  Firefox Internet Explorer   Opera	Safari	Edge
 * Basic support	41   	34      (No)  	            (No)	10      ?
 * -------------------------------------------------------------------------------
 */

 
function main()
{
	
	var str1 = 'Breaded Mushrooms';
	Session.Output(str1.padStart(25, '.'));
}

main();