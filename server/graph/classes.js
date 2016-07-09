
class EmailAddress {
  constructor(options) {
    this.address = options.address;
    this.name = options.name || options.address;
  }
}

class Recipient {
  constructor(options) {
    this.emailAddress = new EmailAddress(options);
  }
}

class ItemBody {
  constructor(options) {
    this.content = options.content;
    this.contentType = options.contentType || 'Text';
  }
}

class TextBody extends ItemBody {
  constructor(content) {
    super({ content, contentType: 'Text' });
  }
}

class Message {
  constructor(options) {
    this.toRecipients = options.toRecipients;
    this.subject = options.subject;
    this.body = options.body;
  }
}

class BasicEmailMessage extends Message {
  constructor(options) {
    super({
      toRecipients: [new Recipient(options.to)],
      subject: options.subject,
      body: new TextBody(options.body),
    });
  }
}

Microsoft.graph = {
  ...Microsoft.graph,
  EmailAddress,
  Recipient,
  ItemBody,
  TextBody,
  Message,
  BasicEmailMessage,
};
