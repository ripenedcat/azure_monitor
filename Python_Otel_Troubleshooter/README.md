# Python OpenTelemetry Troubleshooter for Azure Application Insights

## Features
- Set to debug logging, which also enables debug SDK log.

- Collect all OTEL, Azure and AppInsights related environment variables.

- Collect information for all installed packages.

- Check below features
  - Cloud_RoleName/Instance
  - Excluded Url
  - Sampling
  - Processor
  - AAD Authentication

![image](https://github.com/ripenedcat/azure_monitor/assets/43979954/2b5d372d-0d88-4b74-bc9c-8ccc1597f8d7)


## Usage
1. Download `troubleshooter.py` and put it into your project directory
2. Please import this file before calling `configure_azure_monitor`, as the import of this file itself sets the logging mode to debug. 
``` python
import troubleshooter
```
3. Run `troubleshoot()` function below `configure_azure_monitor`. If you are using Azure AD Authentication, optionally pass the ClientSecretCrendential object to it. 
```
import troubleshooter

configure_azure_monitor(connection_string=AIConnectionString,credential=credential)
troubleshooter.troubleshoot(credential)
```
4. For checking Processor issues, implement `troubleshooter.check_processor(span)` inside the `on_end()` function of your custom processor, as below
```python
class SpanEnrichingProcessor(SpanProcessor):
    def on_end(self, span):
        span._name = "Updated-" + span.name
        span._attributes["CustomDimension1"] = "on_end_1"
        span._attributes["CustomDimension2"] = "on_end_2"
        span._attributes["http.client_ip"] = "1.1.1.1"
        span._attributes["enduser.id"] = "<User ID>"
        troubleshooter.check_processor(span)
```
