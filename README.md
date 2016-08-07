# Microsoft meteor package
__An implementation of the Microsoft OAuth flow using the v2.0 authentication endpoint preview__

[![Build Status][travis-image]][travis-url]

## Getting started

This package is experimental and is not on Atmosphere (the original q42 package is). Add the package to meteor by cloning into the `packages` directory in your Meteor app and running.
```
meteor add ghobbs:microsoft2
```

## Why v2.0?

v2.0 is in preview right now, and it only works with a [handful of APIs](https://msdn.microsoft.com/en-us/office/office365/howto/authenticate-office-365-apis-using-v2). However, the v2.0 API works with the super-useful Graph and Outlook APIs, and provides a single interface for servicing work, school and personal Microsoft accounts. If you want to plug into Outlook and reach the most users, the v2.0 API might be your best bet, even though it is in preview (see references).

## Basic usage

The usage is pretty much the same as all other OAuth flow implementations for meteor. It's almost completely copied from Q42's package for authenticating with the v1.0 API, which was inspired by the official Google meteor package.
Basically you can use:

```javascript
var callback = Accounts.oauth.credentialRequestCompleteHandler(callback);
Microsoft.requestCredential(options, callback);
```


## References

### Original Microsoft endpoint package

* [q42:microsoft](https://github.com/Q42/meteor-microsoft)

### Microsoft REST documentation

* [v2.0 authentication endpoint](https://msdn.microsoft.com/en-us/office/office365/howto/authenticate-office-365-apis-using-v2)
* [More information on available scopes](https://azure.microsoft.com/en-us/documentation/articles/active-directory-v2-scopes/)
* [Getting Started with Mail, Calendar, and Contacts REST APIs](https://dev.outlook.com/restapi/getstarted)

[travis-url]: https://travis-ci.org/gwhobbs/meteor-microsoft
[travis-image]: http://img.shields.io/travis/gwhobbs/meteor-microsoft.svg
