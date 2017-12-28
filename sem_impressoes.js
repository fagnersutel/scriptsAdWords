function main() {

// Enter your account name and email here:

var accountName = “nome da conta";

var yourEmail = “seu email”;

var emailBody=”Ontem estas campanhas não tiveram impressões:<br>”;

var noImpCamp=0;

var campaignsIterator = AdWordsApp.campaigns().get();

while (campaignsIterator.hasNext()) {

var campaign = campaignsIterator.next();

var stats = campaign.getStatsFor(‘YESTERDAY’);

if ((stats.getImpressions() == 0)&&(campaign.isEnabled())) {

emailBody = emailBody + campaign.getName() + “<br>”;

noImpCamp++;

}

}

if (noImpCamp > 0) {

MailApp.sendEmail(yourEmail,”Alerta: ” + accountName + ” – ” + noImpCamp + ” Campanhas sem impressões”,””,{htmlBody: emailBody})

}

}