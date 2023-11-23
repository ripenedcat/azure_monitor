# Python OpenTelemetry Troubleshooter for Azure Application Insights

## Features
- Set to debug logging, which also enables debug SDK log.

- Collect all OTEL and AppInsights related environment variables.

- Check below features
  - cloud_RoleName/Instance
  - Excluded Url
  - Sampling
  - AAD Auth
  - Processor

![image](https://github.com/ripenedcat/azure_monitor/assets/43979954/2b5d372d-0d88-4b74-bc9c-8ccc1597f8d7)


## Usage
1. Download troubleshooter.py and put it into your project directory
2. Import this file
``` python
import troubleshooter
```
3. Run `troubleshoot()` function below `configure_azure_monitor`
```
configure_azure_monitor(credential=credential)
troubleshooter.troubleshoot(credential)
```
4. For checking Processor, implement `troubleshooter.check_processor(span)` as below
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
