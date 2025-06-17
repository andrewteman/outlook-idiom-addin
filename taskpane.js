const idioms = [
  "You can't butter a squirrel with soup in the mailbox.",
  "You can't trust a reindeer with pinecones in the cupboard.",
  "You can't cook a goose with shrimp sandwich in the bog.",
  "You can't drink coffee with a moose in the ice.",
  "You can't take a crap on a cat with porridge in the Malmö.",
  "The sauna won't heat itself with pinecones.",
  "As lost as a moose in Malmö.",
  "He drinks coffee with a helmet on.",
  "Don't trust a squirrel in slippers.",
  "You can't butter both sides of the elk.",
  "You look like you sold the butter and then lost the money.",
  "Everyone knows the monkey, but the monkey knows no-one.",
  "I sense owls in the bog.",
  "Sliding in on a shrimp sandwich.",
  "Like a cat around hot porridge.",
  "Having an unplucked goose with someone.",
  "Cooking soup on a nail.",
  "There's no cow on the ice.",
  "Getting caught with the beard in the mailbox.",
  "You just took a crap in the blue cupboard.",
  "The goose doesn't honk for free."
];

function addIdiom() {
  const item = Office.context.mailbox.item;
  const idiom = idioms[Math.floor(Math.random() * idioms.length)];
  const idiomLine = `As the Swedes often say, ${idiom}`;

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      let body = result.value;
      const signOffIndex = body.search(/(Best|Regards|Sincerely|Cheers|Thanks),/i);
      if (signOffIndex !== -1) {
        body = body.slice(0, signOffIndex) + idiomLine + "\n" + body.slice(signOffIndex);
      } else {
        body += "\n" + idiomLine;
      }
      item.body.setAsync(body, { coercionType: Office.CoercionType.Text });
    }
  });
}
