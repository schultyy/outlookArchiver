DIRECTORY_ROOT = "C:\\temp\\emailRoot\\";

var mailRepository = new MailRepository();
var fs = require("fs");

console.log("Reading emails");

var personalFolders = mailRepository.getChildrenFor("Personal Folders");

var inboxFolder = null;

for (var i = 0; i < personalFolders.length; i++) {
    if (personalFolders[i].Name == "Inbox") {
        inboxFolder = personalFolders[i];
        break;
    }
}

console.assert(inboxFolder != null);

var emails = mailRepository.getMailsForFolder(inboxFolder.UniqueId);

console.log("got " + emails.length + " mails");

if (!fs.exists(DIRECTORY_ROOT))
    fs.createDirectory(DIRECTORY_ROOT);

for (var i = 0; i < emails.length; i++) {
    var currentMail = emails[i];

    if (currentMail.Attachments.length == 0)
        continue;

    console.log("Mail: " + currentMail.Subject);
    console.log("Attachment count: " + currentMail.Attachments.length);

    console.assert(currentMail.SenderName != null && currentMail.SenderName != "");

    var fullPath = fs.combine(DIRECTORY_ROOT, "mails", currentMail.SenderName);

    if (!fs.exists(fullPath)) {
        fs.createDirectory(fullPath);
    }

    for (var x = 0; x < currentMail.Attachments.length; x++) {

        var currentAttachment = currentMail.Attachments[x];
        console.log("filename: " + currentAttachment.Filename);

        var attachmentFilename = fs.combine(fullPath, currentAttachment.Filename);
        mailRepository.saveAttachment(currentMail.UniqueId, currentAttachment, attachmentFilename);

        console.log("Saved attachment: " + attachmentFilename);

        mailRepository.deleteAttachment(currentMail.UniqueId, currentAttachment);

        console.log("Deleted attachment: " + currentAttachment.Filename);
    }
}