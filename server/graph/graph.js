Microsoft.graph = {
  baseUrl: 'https://graph.microsoft.com/v1.0/me/',
  call(method, relativeUrl, userId, options) {
    const url = Microsoft.graph.baseUrl + relativeUrl;
    return Microsoft.http.callAsUserId(method, url, userId, options);
  },
  listMessages(userId, options) {
    return Microsoft.graph.call('GET', 'messages', userId, options);
  },
  sendMail(message, userId) {
    var data = { Message: message };
    return Microsoft.graph.call('POST', 'sendMail', userId, { data });
  },
  sendBasicMail({ toName, toAddress, subject, body }, userId) {
    const message = new BasicEmailMessage({
      body,
      subject,
      to: {
        name: toName,
        address: toAddress,
      },
    });
    return Microsoft.graph.sendMail(message, userId);
  }
};
