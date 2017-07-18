# Exchange.WebServices.OccurrenceData 

![](https://ci.appveyor.com/api/projects/status/i7hsmv33lknd9871?svg=true)

Provide a tool who calculate occurrences of **Exchange** recurring master appointment.

## How to use

```cs
var collection = await OccurrenceCollection.Bind(service, appointment.Id);
```