import logging, subprocess
from os import environ

from opentelemetry.semconv.resource import ResourceAttributes
from opentelemetry.trace import get_tracer_provider
from fixedint import Int32

OTEL_RESOURCE_ATTRIBUTES = "OTEL_RESOURCE_ATTRIBUTES"
OTEL_SERVICE_NAME = "OTEL_SERVICE_NAME"
SERVICE_NAME = ResourceAttributes.SERVICE_NAME

_HASH = 5381
_INTEGER_MAX: int = Int32.maxval
_INTEGER_MIN: int = Int32.minval

logging.basicConfig(format="%(asctime)s:%(levelname)s:%(message)s", level=logging.DEBUG)
logging.info(f"Logging level has been set to debug.")


def get_cloud_rolename():
    tracer_provider = get_tracer_provider()
    resource = tracer_provider.resource  # type: ignore
    map = {}
    map[ResourceAttributes.SERVICE_NAME] = resource.attributes.get(ResourceAttributes.SERVICE_NAME)
    map[ResourceAttributes.SERVICE_NAMESPACE] = resource.attributes.get(ResourceAttributes.SERVICE_NAMESPACE)
    map[ResourceAttributes.SERVICE_INSTANCE_ID] = resource.attributes.get(ResourceAttributes.SERVICE_INSTANCE_ID)
    logging.info(f"The env_resource_map at exporter side is {map}")

def get_excluded_url():
    _root =  r"OTEL_PYTHON_{}"
    excluded_urls = environ.get(_root.format("EXCLUDED_URLS"), "")
    if excluded_urls:
        excluded_url_list = [
            excluded_url.strip() for excluded_url in excluded_urls.split(",")
        ]
    else:
        excluded_url_list = []
    logging.info(f"Parsed exlucded url list is: {excluded_url_list}")

def get_all_env():
    logging.info(f"Listing all environment variables related to Azure, OTEL and ApplicationInsights.")
    env_sub_key_list = ["applicationinsights","otel","azure"]
    for key, value in environ.items():
        if any(sub_key in key.lower() for sub_key in env_sub_key_list):
            logging.info("{0}: {1}".format(key, value))

def check_azure_ad_auth(credential):
    if credential is None:
        logging.info(f"There is no credential passed to the troubleshooter. Skip checking Entra ID Authentication.")
        return
    try:
        token = credential.get_token("https://monitor.azure.com//.default", enable_cae=False)
        logging.info(f"Entra ID Auth is valid. token = {token}")
    except Exception as e:
        logging.error(f"Entra ID Auth is invalid.")
        logging.error(str(e))

def _get_DJB2_sample_score(trace_id_hex) -> float:
    # This algorithm uses 32bit integers
    hash_value = Int32(_HASH)
    for char in trace_id_hex:
        hash_value = ((hash_value << 5) + hash_value) + ord(char)

    if hash_value == _INTEGER_MIN:
        hash_value = int(_INTEGER_MAX)
    else:
        hash_value = abs(hash_value)

    # divide by _INTEGER_MAX for value between 0 and 1 for sampling score
    return float(hash_value) / _INTEGER_MAX

def check_processor(span = None):
    if not span:
        return
    trace_id = span.context.trace_id
    operation_id = "{:032x}".format(trace_id)
    sampling_score = _get_DJB2_sample_score(operation_id)
    span_events = [event.name for event in span.events]
    logging.info(f"Current Span: name = {span.name}, \n has below attributes: {span.attributes},\n  has below events: {span_events}. \n"
                 f"operation_id = {operation_id}. This operation_id represents a sampling score of {sampling_score}, which is smaller than current sampling ratio: {environ.get("OTEL_TRACES_SAMPLER_ARG", "1.0")}, will be sampled and recorded.")


def get_all_packages():
    logging.info("Current Installed Packages using pip list:")
    installed_packages = subprocess.run(['pip', 'list'], stdout=subprocess.PIPE).stdout.decode('utf-8')
    logging.info(f"\n{installed_packages}")


def troubleshoot(credential=None):
    logging.info(f"Starting Python Otel SDK troubleshooter.")
    try:
        get_all_packages()
        get_all_env()
        get_cloud_rolename()
        get_excluded_url()
        check_azure_ad_auth(credential)
    except Exception as e:
        logging.error(f"The troubleshooter encountered an error.")
        logging.error(str(e))




