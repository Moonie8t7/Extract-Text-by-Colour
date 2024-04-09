/**
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @file https://www.reddit.com/r/GoogleAppsScript/comments/1bzukqu/google_sheets_extract_text_by_color_script/
 * @desc Filter values in a range based on font colour and returns a list of matching values
 * @license MIT
 * @version 1.0
 */

/**
 * Filters values in a given range based on the font color and returns matching values.
 * @param {string} color The font color to filter by, specified in hexadecimal format (e.g., "#RRGGBB").
 * @param {Range} range The range of cells to search for matching font colors.
 * @param {number} startcol The starting column index (1-based) of the range.
 * @param {number} startrow The starting row index (1-based) of the range.
 * @return {Array} An array containing values from the specified range where the font color matches the specified color.
 * @customfunction
 */
function LISTIFCOLOUR(color, range, startcol, startrow) {
  try {
    if (typeof color !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(color)) {
      throw new Error(
        'Invalid color format. Please provide color in hexadecimal format (e.g., "#RRGGBB").'
      );
    }
    if (
      !Array.isArray(range) ||
      range.length === 0 ||
      !Array.isArray(range[0])
    ) {
      throw new Error(
        "Invalid range. Please provide a valid two-dimensional array range."
      );
    }
    if (
      typeof startcol !== "number" ||
      typeof startrow !== "number" ||
      startcol < 1 ||
      startrow < 1
    ) {
      throw new Error(
        "Invalid startcol or startrow. Please provide positive integer values."
      );
    }
    // Get the Spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Convert from Integer to Alphanumeric
    const startColId = String.fromCharCode(64 + startcol);
    const endColId = String.fromCharCode(64 + startcol + range[0].length - 1);
    // Set the end row
    const endRow = startrow + range.length - 1;
    // Convert the range to a String
    const rangeString = `${startColId + startrow}:${endColId}${endRow}`;
    // Get the font colours for the range
    const fontColors = ss.getRange(rangeString).getFontColorObjects();
    // Get the text values for the range
    const values = ss.getRange(rangeString).getValues();
    // Empty Array to store the results
    const result = [];

    // Filtering based on font color
    for (let i = 0; i < fontColors.length; i++) {
      for (let j = 0; j < fontColors[i].length; j++) {
        // Saftey net to colour check, in the rare case that Google returns a "named" colour instead
        if (
          fontColors[i][j] &&
          colourNameToHex_(fontColors[i][j].asRgbColor().asHexString()) ===
            color
        ) {
          result.push(values[i][j]);
        }
      }
    }
    return result;
  } catch (error) {
    console.error("Error in LISTIFCOLOUR function:", error);
    throw new Error(
      "An error occurred while processing the function. Please contact support."
    );
  }
}

/**
 * Converts a color name to its corresponding hexadecimal value.
 * @param {string} colour The color name to convert.
 * @return {string} The hexadecimal value of the color, or the original input if no match is found.
 */
function colourNameToHex_(colour) {
  try {
    if (typeof colour !== "string") {
      throw new Error("Input must be a string.");
    }
    const colours = {
      aliceblue: "#f0f8ff",
      antiquewhite: "#faebd7",
      aqua: "#00ffff",
      aquamarine: "#7fffd4",
      azure: "#f0ffff",
      beige: "#f5f5dc",
      bisque: "#ffe4c4",
      black: "#000000",
      blanchedalmond: "#ffebcd",
      blue: "#0000ff",
      blueviolet: "#8a2be2",
      brown: "#a52a2a",
      burlywood: "#deb887",
      cadetblue: "#5f9ea0",
      chartreuse: "#7fff00",
      chocolate: "#d2691e",
      coral: "#ff7f50",
      cornflowerblue: "#6495ed",
      cornsilk: "#fff8dc",
      crimson: "#dc143c",
      cyan: "#00ffff",
      darkblue: "#00008b",
      darkcyan: "#008b8b",
      darkgoldenrod: "#b8860b",
      darkgray: "#a9a9a9",
      darkgreen: "#006400",
      darkkhaki: "#bdb76b",
      darkmagenta: "#8b008b",
      darkolivegreen: "#556b2f",
      darkorange: "#ff8c00",
      darkorchid: "#9932cc",
      darkred: "#8b0000",
      darksalmon: "#e9967a",
      darkseagreen: "#8fbc8f",
      darkslateblue: "#483d8b",
      darkslategray: "#2f4f4f",
      darkturquoise: "#00ced1",
      darkviolet: "#9400d3",
      deeppink: "#ff1493",
      deepskyblue: "#00bfff",
      dimgray: "#696969",
      dodgerblue: "#1e90ff",
      firebrick: "#b22222",
      floralwhite: "#fffaf0",
      forestgreen: "#228b22",
      fuchsia: "#ff00ff",
      gainsboro: "#dcdcdc",
      ghostwhite: "#f8f8ff",
      gold: "#ffd700",
      goldenrod: "#daa520",
      gray: "#808080",
      green: "#008000",
      greenyellow: "#adff2f",
      honeydew: "#f0fff0",
      hotpink: "#ff69b4",
      indianred: "#cd5c5c",
      indigo: "#4b0082",
      ivory: "#fffff0",
      khaki: "#f0e68c",
      lavender: "#e6e6fa",
      lavenderblush: "#fff0f5",
      lawngreen: "#7cfc00",
      lemonchiffon: "#fffacd",
      lightblue: "#add8e6",
      lightcoral: "#f08080",
      lightcyan: "#e0ffff",
      lightgoldenrodyellow: "#fafad2",
      lightgrey: "#d3d3d3",
      lightgreen: "#90ee90",
      lightpink: "#ffb6c1",
      lightsalmon: "#ffa07a",
      lightseagreen: "#20b2aa",
      lightskyblue: "#87cefa",
      lightslategray: "#778899",
      lightsteelblue: "#b0c4de",
      lightyellow: "#ffffe0",
      lime: "#00ff00",
      limegreen: "#32cd32",
      linen: "#faf0e6",
      magenta: "#ff00ff",
      maroon: "#800000",
      mediumaquamarine: "#66cdaa",
      mediumblue: "#0000cd",
      mediumorchid: "#ba55d3",
      mediumpurple: "#9370d8",
      mediumseagreen: "#3cb371",
      mediumslateblue: "#7b68ee",
      mediumspringgreen: "#00fa9a",
      mediumturquoise: "#48d1cc",
      mediumvioletred: "#c71585",
      midnightblue: "#191970",
      mintcream: "#f5fffa",
      mistyrose: "#ffe4e1",
      moccasin: "#ffe4b5",
      navajowhite: "#ffdead",
      navy: "#000080",
      oldlace: "#fdf5e6",
      olive: "#808000",
      olivedrab: "#6b8e23",
      orange: "#ffa500",
      orangered: "#ff4500",
      orchid: "#da70d6",
      palegoldenrod: "#eee8aa",
      palegreen: "#98fb98",
      paleturquoise: "#afeeee",
      palevioletred: "#d87093",
      papayawhip: "#ffefd5",
      peachpuff: "#ffdab9",
      peru: "#cd853f",
      pink: "#ffc0cb",
      plum: "#dda0dd",
      powderblue: "#b0e0e6",
      purple: "#800080",
      red: "#ff0000",
      rosybrown: "#bc8f8f",
      royalblue: "#4169e1",
      saddlebrown: "#8b4513",
      salmon: "#fa8072",
      sandybrown: "#f4a460",
      seagreen: "#2e8b57",
      seashell: "#fff5ee",
      sienna: "#a0522d",
      silver: "#c0c0c0",
      skyblue: "#87ceeb",
      slateblue: "#6a5acd",
      slategray: "#708090",
      snow: "#fffafa",
      springgreen: "#00ff7f",
      steelblue: "#4682b4",
      tan: "#d2b48c",
      teal: "#008080",
      thistle: "#d8bfd8",
      tomato: "#ff6347",
      turquoise: "#40e0d0",
      violet: "#ee82ee",
      wheat: "#f5deb3",
      white: "#ffffff",
      whitesmoke: "#f5f5f5",
      yellow: "#ffff00",
      yellowgreen: "#9acd32"
    };
    const lowerCaseColour = colour.toLowerCase();
    if (colours.hasOwnProperty(lowerCaseColour)) {
      return colours[lowerCaseColour];
    }
    return colour;
  } catch (error) {
    console.error("Error in colourNameToHex_ function:", error);
    throw new Error(
      "An error occurred while processing the function. Please contact support."
    );
  }
}
