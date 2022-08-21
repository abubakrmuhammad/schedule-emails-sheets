/**
 * @OnlyCurrentDoc
 */

const FIRST_NAME_COL_NAME = 'First Name';
const LAST_NAME_COL_NAME = 'Last Name';
const RECIPIENT_EMAIL_COL_NAME = 'Email';
const EMAIL_STATUS_COL_NAME = 'Email Status';
const SCHEDULE_SENT_COL_NAME = 'Scheduled/Sent Date/Time';
const SCHEDULE_DATA_COL_NAME = `Schedule Data (Don't Touch)`;

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Mail Merge')
    .addItem('Send Emails', sendEmails.name)
    .addItem('Schedule Emails', scheduleEmails.name)
    .addToUi();
}

function sendEmails() {
  try {
    const controller = new Controller();

    controller.init();
  } catch (e) {
    throw new Error(e);
  }
}

function scheduleEmails() {
  try {
    const controller = new ScheduledController();

    controller.init();
  } catch (e) {
    throw new Error(e);
  }
}

function sendScheduledEmail() {
  const controller = new ScheduledController();

  controller.sendScheduledEmail();
}

class Controller {
  protected sheet: GoogleAppsScript.Spreadsheet.Sheet;
  protected gmail: GoogleAppsScript.Gmail.GmailApp;

  protected subjectline: string;
  protected sheetData: SheetData;
  protected draftTemplate: DraftTemplate;

  protected parsedData: ParsedRow[];
  protected rowsToUse: ParsedRow[];

  protected columnNumbers: ColumnNumbers;

  constructor() {
    this.sheet = SpreadsheetApp.getActiveSheet();
    this.gmail = GmailApp;

    this.sendEmail = this.sendEmail.bind(this);
  }

  init() {
    // Ask for Draft Email Subject
    this.askForDraftSubject();

    // Get the draft from Gmail
    this.getDraftEmail();

    // Get the data from sheet
    this.getDataFromSheet();

    // Parse the data
    this.parseData();

    // Find only the unscheduled Emails
    this.useOnlyUnscheduledEmailRows();

    // Find only unsent Emails
    this.useOnlyUnsentEmailRows();

    // Add the required data to the template
    this.fillInDraftTemplatesFromData();

    // Send Emails
    this.rowsToUse.forEach(this.sendEmail);
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

    if (!draftToUse) throw 'No Gmail draft with that subject found';

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

  parseData() {
    const parsedData = this.sheetData.mappedRows.map<ParsedRow>((row) => ({
      rowNumber: parseInt(row.rowIndex) + 1,
      email: row[RECIPIENT_EMAIL_COL_NAME].trim(),
      emailStatus: row[EMAIL_STATUS_COL_NAME].trim(),
      isSent: row[EMAIL_STATUS_COL_NAME].trim() === EmailStatus.Sent,
      isScheduled: !!row[SCHEDULE_SENT_COL_NAME].trim(),
      scheduledDateTime: new Date(row[SCHEDULE_SENT_COL_NAME].trim()),
      scheduleData: row[SCHEDULE_DATA_COL_NAME],
      hasScheduleData: !!row[SCHEDULE_DATA_COL_NAME],
      filledTemplate: this.draftTemplate,
    }));

    this.columnNumbers = {
      firstName: this.sheetData.headerRow.indexOf(FIRST_NAME_COL_NAME) + 1,
      lastName: this.sheetData.headerRow.indexOf(LAST_NAME_COL_NAME) + 1,
      email: this.sheetData.headerRow.indexOf(RECIPIENT_EMAIL_COL_NAME) + 1,
      emailStatus: this.sheetData.headerRow.indexOf(EMAIL_STATUS_COL_NAME) + 1,
      scheduleOrSent:
        this.sheetData.headerRow.indexOf(SCHEDULE_SENT_COL_NAME) + 1,
      scheduleData: this.sheetData.headerRow.indexOf(SCHEDULE_DATA_COL_NAME) + 1,
    };

    this.parsedData = parsedData;
    this.rowsToUse = parsedData;
  }

  useOnlyUnscheduledEmailRows() {
    this.rowsToUse = this.rowsToUse.filter((row) => !row.isScheduled);
  }

  useOnlyScheduledEmailRows() {
    this.rowsToUse = this.rowsToUse.filter((row) => row.isScheduled);
  }

  useOnlyUnsentEmailRows() {
    this.rowsToUse = this.rowsToUse.filter((row) => !row.isSent);
  }

  sendEmail(row: ParsedRow) {
    const templateData = row.filledTemplate;

    this.gmail.sendEmail(row.email, templateData.subject, templateData.textBody, {
      htmlBody: templateData.htmlBody,
      attachments: templateData.attachments as any,
    });

    this.sheet
      .getRange(row.rowNumber, this.columnNumbers.emailStatus)
      .setValue(EmailStatus.Sent);

    this.sheet
      .getRange(row.rowNumber, this.columnNumbers.scheduleOrSent)
      .setValue(new Date());
  }

  protected fillInDraftTemplatesFromData() {
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

    this.rowsToUse.forEach((row) => {
      const templateString = JSON.stringify(row.filledTemplate);
      const mappedRow = this.sheetData.mappedRows.find(
        (mappedRow) => mappedRow[RECIPIENT_EMAIL_COL_NAME] === row.email
      );

      const filledTemplateString = templateString.replace(
        /{([^{}]+)}/g,
        (_, key) => escapeData(mappedRow[key] || '')
      );

      row.filledTemplate = JSON.parse(filledTemplateString);
    });
  }
}

class ScheduledController extends Controller {
  constructor() {
    super();

    this.verifySchedules = this.verifySchedules.bind(this);
  }

  public init() {
    // Clear All previous schedules
    this.clearExistingSchedules();

    // Ask for Draft Email Subject
    this.askForDraftSubject();

    // Get the draft from Gmail
    this.getDraftEmail();

    // Get the data from sheet
    this.getDataFromSheet();

    // Parse the data
    this.parseData();

    // Use only scheduled Emails
    this.useOnlyScheduledEmailRows();

    // Use only unsent Emails
    this.useOnlyUnsentEmailRows();

    // Use emails with a valid schedule date
    this.verifySchedules();

    // Add the required data to the template
    this.fillInDraftTemplatesFromData();

    // Add the template to the sheet
    this.addTemplateDataToSheet();

    // Schedule Emails
    this.createScheduleTriggers();
  }

  public sendScheduledEmail() {
    // Get the data from sheet
    this.getDataFromSheet();

    // Parse the data
    this.parseData();

    // Use the rows that has schedule data
    this.useRowsWithScheduleDataOnly();

    // Get draft data from sheet
    this.getTemplateDataFromSheet();

    // Find the row for current trigger
    const rowToUse = this.rowsToUse.find(
      (row) => row.scheduledDateTime.getTime() <= new Date().getTime()
    );

    // Send the Email
    this.sendEmail(rowToUse);
  }

  protected clearExistingSchedules() {
    const allTriggers = ScriptApp.getProjectTriggers();

    for (const trigger of allTriggers) {
      if (trigger.getHandlerFunction() === sendScheduledEmail.name)
        ScriptApp.deleteTrigger(trigger);
    }
  }

  protected verifySchedules() {
    const now = new Date();

    this.rowsToUse.forEach((row) => {
      const isValid = row.scheduledDateTime > now;

      if (!isValid)
        this.sheet
          .getRange(row.rowNumber, this.columnNumbers.emailStatus)
          .setValue(EmailStatus.InvalidSchedule);
    });

    this.rowsToUse = this.rowsToUse.filter((row) => row.scheduledDateTime > now);
  }

  protected addTemplateDataToSheet() {
    const draftTemplateString = JSON.stringify(this.draftTemplate);

    this.rowsToUse.forEach((row, i) => {
      this.sheet
        .getRange(row.rowNumber, this.columnNumbers.scheduleData)
        .setValue(draftTemplateString);
    });
  }

  protected getTemplateDataFromSheet() {
    this.rowsToUse = this.rowsToUse.map((row) => {
      const templateData = this.sheet
        .getRange(row.rowNumber, this.columnNumbers.scheduleData)
        .getValue();

      return {
        ...row,
        filledTemplate: JSON.parse(templateData),
      };
    });
  }

  protected createScheduleTriggers() {
    this.rowsToUse.forEach((row) => {
      const triggerDate = row.scheduledDateTime;

      ScriptApp.newTrigger(sendScheduledEmail.name)
        .timeBased()
        .at(triggerDate)
        .inTimezone(
          SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
        )
        .create();

      this.sheet
        .getRange(row.rowNumber, this.columnNumbers.emailStatus)
        .setValue(EmailStatus.Scheduled);
    });
  }

  useRowsWithScheduleDataOnly() {
    this.rowsToUse = this.rowsToUse.filter(
      (row) => row.emailStatus === EmailStatus.Scheduled
    );
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
    const mappedRows = rows.map((row, rowIndex) => {
      const mappedRow: MappedRow = {
        rowIndex: (rowIndex + 1).toString(),
      };

      this.headerRow.forEach((heading, i) => {
        mappedRow[heading] = row[i];
      });

      return mappedRow;
    });

    this.mappedRows = mappedRows;
  }
}

type MappedRow = {
  [key: string]: string;
  rowIndex: string;
};

type ParsedRow = {
  rowNumber: number;
  email: string;
  scheduledDateTime: Date;
  emailStatus: string;
  isScheduled: boolean;
  isSent: boolean;
  scheduleData: string;
  hasScheduleData: boolean;
  filledTemplate: DraftTemplate;
};

type DraftTemplate = {
  subject: string;
  textBody: string;
  htmlBody: string;
  attachments: GoogleAppsScript.Gmail.GmailAttachment[];
};

enum EmailStatus {
  Sent = 'Sent',
  Scheduled = 'Scheduled',
  InvalidSchedule = 'Invalid Schedule Date/Time',
}

interface ColumnNumbers {
  firstName: number;
  lastName: number;
  email: number;
  emailStatus: number;
  scheduleOrSent: number;
  scheduleData: number;
}
