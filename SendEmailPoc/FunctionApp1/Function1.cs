using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.SendMail;

public class Function1
{
  private readonly ILogger<Function1> _logger;
  private readonly GraphServiceClient _graphServiceClient;

  public Function1(ILogger<Function1> logger)
  {
    _logger = logger;
    var managedIdentityClientId = Environment.GetEnvironmentVariable("UserManagedIdentityProviderClientId");
    var credential = new ManagedIdentityCredential(managedIdentityClientId);
    _graphServiceClient = new GraphServiceClient(credential);
  }

  [Function("Function1")]
  public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req, CancellationToken ct)
  {
    try
    {
      var sender = req.Query["sender"].FirstOrDefault() ?? Environment.GetEnvironmentVariable("EMAIL_DEFAULT_SENDER");
      var to = req.Query["to"].FirstOrDefault() ?? Environment.GetEnvironmentVariable("EMAIL_DEFAULT_RECIPIENT");
      var subject = req.Query["subject"].FirstOrDefault() ?? Environment.GetEnvironmentVariable("EMAIL_DEFAULT_SUBJECT");
      var bodyText = req.Query["body"].FirstOrDefault() ?? Environment.GetEnvironmentVariable("EMAIL_DEFAULT_BODY");

      var user = await _graphServiceClient.Users[sender].GetAsync(cancellationToken: ct);
      if (user == null)
      {
        return new OkObjectResult($"Mail NOT sent: sender '{sender}' was not found in Microsoft Graph.");
      }

      var message = new Message
      {
        Subject = subject,
        Body = new ItemBody { ContentType = BodyType.Text, Content = bodyText },
        ToRecipients = to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(addr => new Recipient { EmailAddress = new EmailAddress { Address = addr } }).ToList()
      };
      await _graphServiceClient.Users[sender].SendMail.PostAsync(new SendMailPostRequestBody { Message = message, SaveToSentItems = true }, cancellationToken: ct);

      return new OkObjectResult($"Mail sent successfully from '{sender}' to '{to}'. {bodyText}");

    }
    catch (Exception ex)
    {
      _logger.LogError(ex, "Failed to send mail or retrieve user.");
      return new StatusCodeResult(StatusCodes.Status500InternalServerError);
    }
  }
}