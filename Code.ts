/**
 * @OnlyCurrentDoc
 */

const FIRST_NAME_COL_NAME = 'First Name';
const LAST_NAME_COL_NAME = 'Last Name';
const RECIPIENT_EMAIL_COL_NAME = 'Email';
const EMAIL_SENT_COL_NAME = 'Email Sent';
const SCHEDULE_COL_NAME = 'Scheduled Date/Time';

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Mail Merge')
    .addItem('Send Emails', 'sendEmails')
    .addItem('Schedule Emails', 'scheduleEmails')
    .addToUi();
}

function sendEmails() {
  try {
    const controller = new Controller();

    // Ask for Draft Email Subject
    controller.askForDraftSubject();

    // Get the draft from gmail
    controller.getDraftEmail();

    // Get the data from sheet
    controller.getDataFromSheet();

    // Parse the data
    const parsedData = controller.getParsedData();

    // Find only the unscheduled Emails
    const unscheduledEmailRows = parsedData.filter((row) => !row.isScheduled);

    // Find only unsent Emails
    const unsentEmailRows = unscheduledEmailRows.filter((row) => !row.isSent);

    // Send Emails
    unsentEmailRows.forEach(controller.sendEmail);
  } catch (e) {
    throw new Error(e);
  }
}

function scheduleEmails() {}

class Controller {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private gmail: GoogleAppsScript.Gmail.GmailApp;

  private subjectline: string;
  private sheetData: SheetData;
  private draftTemplate: DraftTemplate;

  constructor() {
    this.sheet = SpreadsheetApp.getActiveSheet();
    this.gmail = GmailApp;

    this.sendEmail = this.sendEmail.bind(this);
  }

  askForDraftSubject(): void {
    const subjectLine: string = Browser.inputBox(
      'Mail Merge',
      'Type or copy/paste the subject line of the Gmail draft message',
      Browser.Buttons.OK_CANCEL
    );

    if (subjectLine === 'cancel' || subjectLine === '')
      throw 'Please provide a subject line';

    this.subjectline = subjectLine;
  }

  getDraftEmail() {
    const drafts = this.gmail.getDrafts();

    const draftToUse = drafts.find(
      (draft) => draft.getMessage().getSubject() === this.subjectline
    );

    if (!draftToUse) throw 'No gmail draft with that subject found';

    const message = draftToUse.getMessage();

    this.draftTemplate = {
      subject: message.getSubject(),
      textBody: message.getPlainBody(),
      htmlBody: message.getBody(),
      attachments: message.getAttachments(),
    };
  }

  getDataFromSheet() {
    const dataRange = this.sheet.getDataRange();
    const data = dataRange.getDisplayValues();

    this.sheetData = new SheetData(data);
  }

  getParsedData(): ParsedRowData[] {
    return this.sheetData.mappedRows.map<ParsedRowData>((row) => ({
      firstName: row[FIRST_NAME_COL_NAME].trim(),
      lastName: row[LAST_NAME_COL_NAME].trim(),
      email: row[RECIPIENT_EMAIL_COL_NAME].trim(),
      emailStatus: row[EMAIL_SENT_COL_NAME].trim(),
      isSent: !!row[EMAIL_SENT_COL_NAME].trim(),
      isScheduled: !!row[SCHEDULE_COL_NAME].trim(),
      scheduledDateTime: new Date(row[SCHEDULE_COL_NAME].trim()),
    }));
  }

  sendEmail(data: ParsedRowData) {
    const templateData = this.fillInDraftTemplateFromData(data);

    this.gmail.sendEmail(
      data.email,
      templateData.subject,
      templateData.textBody,
      {
        htmlBody: templateData.htmlBody,
        attachments: templateData.attachments as any,
      }
    );
  }

  private fillInDraftTemplateFromData(data: ParsedRowData): DraftTemplate {
    const escapeData = (str) =>
      str
        .replace(/[\\]/g, '\\\\')
        .replace(/[\"]/g, '\\"')
        .replace(/[\/]/g, '\\/')
        .replace(/[\b]/g, '\\b')
        .replace(/[\f]/g, '\\f')
        .replace(/[\n]/g, '\\n')
        .replace(/[\r]/g, '\\r')
        .replace(/[\t]/g, '\\t');

    const templateString = JSON.stringify(this.draftTemplate);

    const filledTemplateString = templateString.replace(/{{[^{}]+}}/g, (key) =>
      escapeData(data[key.replace(/[{}]+/g, '')] || '')
    );

    return JSON.parse(filledTemplateString);
  }
}

class SheetData {
  public headerRow: string[];
  public mappedRows: MappedRow[];

  constructor(data: string[][]) {
    this.headerRow = data.shift();

    this.convertRowsToMappedRows(data);
  }

  private convertRowsToMappedRows(rows: string[][]) {
    const mappedRows = rows.map((row) => {
      const column: MappedRow = {};

      this.headerRow.forEach((heading, i) => {
        column[heading] = row[i];
      });

      return column;
    });

    this.mappedRows = mappedRows;
  }
}

type MappedRow = {
  [key: string]: string;
};

type ParsedRowData = {
  firstName: string;
  lastName: string;
  email: string;
  scheduledDateTime: Date;
  emailStatus: string;
  isScheduled: boolean;
  isSent: boolean;
};

type DraftTemplate = {
  subject: string;
  textBody: string;
  htmlBody: string;
  attachments: GoogleAppsScript.Gmail.GmailAttachment[];
};
