import { SPCrudOperations } from "./SPCrudOperations";
import { SPHttpClient } from '@microsoft/sp-http';

export class SPHelpers {
  private padWithZero(num: number): string {
    return num < 10 ? `0${num}` : `${num}`;
  }

  public convertGMTToLocalTime24Hour(gmtDate: string): string {
    // Create a Date object from the GMT date string
    const date = new Date(gmtDate); // JavaScript Date will automatically interpret the GMT date
    // Extract hours, minutes, and seconds in the local time zone
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = date.getSeconds();
    // Format the time as HH:mm:ss
    const formattedTime = `${this.padWithZero(hours)}:${this.padWithZero(
      minutes
    )}:${this.padWithZero(seconds)}`;
    return formattedTime;
  }

  public convertGMTToLocalTime12Hour(gmtDate: string): string {
    // Create a Date object from the GMT date string
    const date = new Date(gmtDate);
    // Extract hours, minutes, and seconds in the local time zone
    let hours = date.getHours();
    const minutes = date.getMinutes();
    const seconds = date.getSeconds();
    // Determine AM or PM
    const ampm = hours >= 12 ? "PM" : "AM";
    // Convert to 12-hour format
    hours = hours % 12;
    if (hours === 0) {
      hours = 12; // Handle the case where 0 hours is 12 in 12-hour format
    }
    // Format the time as hh:mm:ss AM/PM
    const formattedTime = `${this.padWithZero(hours)}:${this.padWithZero(
      minutes
    )}:${this.padWithZero(seconds)} ${ampm}`;

    return formattedTime;
  }

  public convertLocalToGMT(localDate: Date): Date {
    // Get the time in milliseconds since January 1, 1970, 00:00:00 UTC
    const utcMilliseconds =
      localDate.getTime() - localDate.getTimezoneOffset() * 60000;

    // Create a new Date object using the UTC milliseconds
    const gmtDate = new Date(utcMilliseconds);

    return gmtDate;
  }

  public setDateWithSelectedTime(date: Date, timeString: string): Date {
    // Split the time string into hours and minutes
    const [hours, minutes] = timeString.split(":").map(Number);

    // Set the hours and minutes on the date object
    date.setHours(hours, minutes, 0, 0); // Setting seconds and milliseconds to 0

    return date;
  }

  public adjustDateForGMTOffset(dateString: string): string {
    const myDate = new Date(dateString);
    const gmtOffset = myDate.getTimezoneOffset(); // Get the offset in minutes
    // Convert minutes to milliseconds
    const offsetInMilliseconds = gmtOffset * 60 * 1000;

    // Add the offset to the date
    const adjustedDate = new Date(myDate.getTime() - offsetInMilliseconds);
    return adjustedDate.toISOString().split("T")[0];
  }

  public CamelString(str: string): string {
    if (!str) return "";
    return str
      .split(" ") // Split into words
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()) // Capitalize each word
      .join(" "); // Join back into a string
  }

  public NormalizeToStringArray(val: any): string[] | undefined {
    if (val === undefined || val === null) return undefined;

    // If already an array, convert all to strings
    if (Array.isArray(val)) {
      return val
        .map((v) => (v !== undefined && v !== null ? String(v) : ''))
        .filter(Boolean);
    }

    // If it's a comma-separated string
    if (typeof val === 'string') {
      return val
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    }

    // If it's a SharePoint REST object with results: []
    if (val && typeof val === 'object' && Array.isArray(val.results)) {
      return val.results.map((v: any) => String(v)).filter(Boolean);
    }

    return undefined;
  }

  public removeWhitespaces(str: string): string {
    return str.replace(/\s+/g, '');
  }

  public getCompanyCode(companyTitle?: string): string {
    if (!companyTitle) return 'UNK';
    const letters = String(companyTitle).replace(/[^A-Za-z]/g, '').toUpperCase();
    return letters.slice(0, 3) || 'UNK';
  }

  public formatYYYYMMDD(d: Date): string {
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}${mm}${dd}`;
  }

  public async generateCoralReferenceNumber(spHttpClient: SPHttpClient, webUrl: string, listTitle: string,
    finalRecord: { Id: number; Created?: string | Date }, companyTitle?: string): Promise<string> {
    const companyCode = this.getCompanyCode(companyTitle);
    const createdDate = finalRecord?.Created
      ? new Date(finalRecord.Created as any)
      : new Date();

    const yyyymmdd = this.formatYYYYMMDD(createdDate);
    const prefix = `${companyCode}-HSE-PPE-${yyyymmdd}-`;
    const esc = (s: string) => s.replace(/'/g, "''");

    // Get the last reference for this date and prefix, then increment NN
    const query = `?$select=Id,CoralReferenceNumber` +
      `&$filter=startswith(CoralReferenceNumber,'${esc(prefix)}')` +
      `&$orderby=CoralReferenceNumber desc` +
      `&$top=1`;

    const reader = new SPCrudOperations(spHttpClient, webUrl, listTitle, query);
    const items: Array<{ Id: number; CoralReferenceNumber?: string }> = await reader._getItemsWithQuery();

    let next = 1;
    if (Array.isArray(items) && items.length) {
      const last = String(items[0]?.CoralReferenceNumber || '');
      const lastNN = last.split('-').pop();
      const parsed = lastNN ? parseInt(lastNN, 10) : NaN;
      if (Number.isFinite(parsed)) next = parsed + 1;
    }
    const nnStr = String(next).padStart(2, '0');
    return `${prefix}${nnStr}`;
  }

  // Convenience: compute and immediately update the item’s CoralReferenceNumber
  public async assignCoralReferenceNumber(spHttpClient: SPHttpClient, webUrl: string, listTitle: string, finalRecord: { Id: number; Created?: string | Date },
    companyTitle?: string
  ): Promise<string> {
    const coralRef = await this.generateCoralReferenceNumber(
      spHttpClient,
      webUrl,
      listTitle,
      finalRecord,
      companyTitle
    );

    const updater = new SPCrudOperations(spHttpClient, webUrl, listTitle, '');
    await updater._updateItem(String(finalRecord.Id), { CoralReferenceNumber: coralRef });
    return coralRef;
  }

  // Read formId from the page URL so the form can be deep-linked
  public getQueryNumber(name: string): number | undefined 
  {
    try {
      const href = (window.top?.location?.href) || window.location.href;
      const v = new URL(href).searchParams.get(name) || undefined;
      const n = v != null ? Number(v) : NaN;
      return Number.isFinite(n) ? n : undefined;
    } catch { return undefined; }
  };

}
