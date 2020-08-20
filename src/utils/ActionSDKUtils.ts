import { ActionError, ActionErrorCode } from "./ActionError";
import * as uuid from "uuid";

export namespace ActionSDKUtils {
  export let YEARS: string = "YEARS";
  export let MONTHS: string = "MONTHS";
  export let WEEKS: string = "WEEKS";
  export let DAYS: string = "DAYS";
  export let HOURS: string = "HOURS";
  export let MINUTES: string = "MINUTES";
  export let DEFAULT_LOCALE: string = "en";

  export function parseUrlQueries(url: string): { [key: string]: string } {
    if (isEmptyString(url) || url.indexOf("?") === -1) {
      return null;
    }
    let params: { [key: string]: string } = {};
    // Separate the queries part
    let queries = url.substr(url.indexOf("?") + 1);
    if (!isEmptyString(queries)) {
      // Decode the queries
      queries = decodeURIComponent(queries);
      // Split the queries to get key-value pairs
      let keyValuePairs = queries.split("&");
      for (let keyValuePair of keyValuePairs) {
        let keyValue = keyValuePair.split("=");
        if (!isEmptyString(keyValue[0]) && !isEmptyString(keyValue[1])) {
          params[keyValue[0]] = keyValue[1];
        }
      }
    }
    return params;
  }

  export function isValidJson(json: string): boolean {
    try {
      JSON.parse(JSON.stringify(json));
      return true;
    } catch (e) {
      return false;
    }
  }

  // To avoid HTML injections, we sanitize all HTML tags
  // by replacing all '<' with '&lt;' and '>' with '&gt;'
  export function sanitizeHtmlTags(str: string): string {
    if (isEmptyString(str)) return str;
    var tagsToReplace = {
      "<": "&lt;",
      ">": "&gt;",
    };
    let sanitizedString = str.replace(/[&<>]/g, (tag) => {
      return tagsToReplace[tag] || tag;
    });
    return sanitizedString;
  }

  export function executeFunction(
    funcNameWithNamespaces: string,
    args: any[] = []
  ) {
    let components = funcNameWithNamespaces.split(".");
    let func;
    for (let i = 0; i < components.length; i++) {
      let component = components[i];
      if (!func) {
        func = window[component];
      } else {
        func = func[component];
      }
    }
    func(args);
  }

  export function replaceCharacterInString(
    str: string,
    oldChar: string,
    newChar: string
  ): string {
    return str.split(oldChar).join(newChar);
  }

  export function jsonIsArray(json: JSON): boolean {
    return Object.prototype.toString.call(json) === "[object Array]";
  }

  export function isEmptyString(str: string): boolean {
    return isEmptyObject(str);
  }

  export function isEmptyObject(obj: any): boolean {
    if (obj == undefined || obj == null) {
      return true;
    }

    var isEmpty = false;

    if (typeof obj === "number" || typeof obj === "boolean") {
      isEmpty = false;
    } else if (typeof obj === "string") {
      isEmpty = obj.trim().length == 0;
    } else if (Array.isArray(obj)) {
      isEmpty = obj.length == 0;
    } else if (typeof obj === "object") {
      if (isValidJson(obj)) {
        isEmpty = JSON.stringify(obj) == "{}";
      }
    }
    return isEmpty;
  }

  export function parseJson(jsonString, defaultValue = null) {
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      return defaultValue || {};
    }
  }

  export function stringifyJson(obj: any) {
    try {
      if (isEmptyObject(obj)) {
        return null;
      }
      return JSON.stringify(obj);
    } catch (e) {
      return null;
    }
  }

  export function getTimeRemaining(deadLineDate: Date): {} {
    var now = new Date().getTime();
    var deadLineTime = deadLineDate.getTime();

    var minutes: number = 0;
    var hours: number = 0;
    var days: number = 0;
    var weeks: number = 0;
    var months: number = 0;
    var years: number = 0;

    var diff = Math.abs(deadLineTime - now);
    if (diff > 0) {
      var minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
      var hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));

      var days = Math.floor(
        (diff % (1000 * 60 * 60 * 24 * 7)) / (1000 * 60 * 60 * 24)
      );
      var weeks = Math.floor(
        (diff % (1000 * 60 * 60 * 24 * 30)) / (1000 * 60 * 60 * 24 * 7)
      );
      var months = Math.floor(
        (diff % (1000 * 60 * 60 * 24 * 365)) / (1000 * 60 * 60 * 24 * 30)
      );
      var years = Math.floor(diff / (1000 * 60 * 60 * 24 * 365));
    }
    return {
      YEARS: years,
      MONTHS: months,
      WEEKS: weeks,
      DAYS: days,
      HOURS: hours,
      MINUTES: minutes,
    };
  }

  export function getDefaultExpiry(activeDays: number): Date {
    let date: Date = new Date();
    date.setDate(date.getDate() + activeDays);

    // round off to next 30 minutes time multiple
    if (date.getMinutes() > 30) {
      date.setMinutes(0);
      date.setHours(date.getHours() + 1);
    } else {
      date.setMinutes(30);
    }
    return date;
  }

  export function isServerURL(url: string): boolean {
    if (!isEmptyString(url) && url.match(/^https?:\/\//)) {
      return true;
    }
    return false;
  }

  export function generateGUID(): string {
    return uuid.v4();
  }

  export function getValues(map: JSON) {
    var values = [];
    for (var key in map) {
      values.push(map[key]);
    }
    return values;
  }

  export function getMaxValue(values: number[]): number {
    let result = Number.MIN_VALUE;
    for (var i = 0; i < values.length; i++) {
      result = Math.max(result, values[i]);
    }
    return result;
  }

  export function downloadContent(fileName: string, data: string) {
    if (data && fileName) {
      var a = document.createElement("a");
      a.href = data;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  }

  export function isRTL(locale: string): boolean {
    let rtlLang: string[] = ["ar", "he", "fl"];
    if (locale && rtlLang.indexOf(locale.split("-")[0]) !== -1) {
      return true;
    } else {
      return false;
    }
  }

  // read a local file's blob Object as ArrayBuffer
  export async function readBlobAsync(
    blob: Blob
  ): Promise<string | ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsArrayBuffer(blob);
      reader.onloadend = function () {
        resolve(reader.result);
      };
      reader.onerror = function (e) {
        let error: ActionError = {
          errorCode: ActionErrorCode.IOException,
          errorMessage: "Error in reading blobUrl: " + e,
        };
        reject(error);
      };
    });
  }
  export function dateTimeToLocaleString(
    date: Date,
    locale: string,
    options?: Intl.DateTimeFormatOptions
  ): string {
    let dateOptions: Intl.DateTimeFormatOptions = options
      ? options
      : {
          year: "numeric",
          month: "long",
          day: "numeric",
          hour: "numeric",
          minute: "numeric",
        };
    return date.toLocaleDateString(
      locale ? locale : DEFAULT_LOCALE,
      dateOptions
    );
  }

  export function announceText(text: string) {
    let ariaLiveSpan: HTMLSpanElement = document.getElementById(
      "aria-live-span"
    );
    if (ariaLiveSpan) {
      ariaLiveSpan.innerText = text;
    } else {
      ariaLiveSpan = document.createElement("SPAN");
      ariaLiveSpan.style.cssText =
        "position: fixed; overflow: hidden; width: 0px; height: 0px;";
      ariaLiveSpan.id = "aria-live-span";
      ariaLiveSpan.innerText = "";
      ariaLiveSpan.setAttribute("aria-live", "polite");
      ariaLiveSpan.tabIndex = -1;
      document.body.appendChild(ariaLiveSpan);
      setTimeout(() => {
        ariaLiveSpan.innerText = text;
      }, 50);
    }
  }

  export function getNonNullString(str: string): string {
    if (isEmptyObject(str)) {
      return "";
    } else {
      return str;
    }
  }
}
