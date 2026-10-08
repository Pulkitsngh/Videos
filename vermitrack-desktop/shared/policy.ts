/** Access policy carried over from the web version of VermiTrack. */
export const KEY_PERSON_EMAIL="pulkit.singh@lnbgroup.com";
export const KEY_PERSON_NAME="Pulkit Singh";
export const DELETION_APPROVER_EMAILS=[KEY_PERSON_EMAIL,"pritam.suryavanshi@lnbgroup.com"];
export const DELETION_APPROVER_NAMES="Pulkit Singh or Pritam Suryavanshi";
export const EDIT_WINDOW_HOURS=24;
export const RECYCLE_BIN_DAYS=90;

export function isDeletionApprover(email:string) { return DELETION_APPROVER_EMAILS.includes(email.trim().toLowerCase()); }
export function isKeyPersonEmail(email:string) { return email.trim().toLowerCase()===KEY_PERSON_EMAIL; }
