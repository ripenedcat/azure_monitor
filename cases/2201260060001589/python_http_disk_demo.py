
import json
import requests
import datetime
import hashlib
import hmac
import base64
import os
import psutil
import time

# Update the customer ID to your Log Analytics workspace ID
customer_id = '<your id here>'

# For the shared key, use either the primary or the secondary Connected Sources client authentication key
shared_key = "<your key here>"

# The log type is the name of the event that is being submitted
log_type = 'Perf'

# An example JSON web monitor object
disk_partitions = psutil.disk_partitions()
system ={
    'Computer': os.uname().nodename,
    'ObjectName': 'Logical Disk',
    'CounterName': 'Free Space',
    'CounterValue': []
}

#####################
######Functions######
#####################

# Build the API signature
def build_signature(customer_id, shared_key, date, content_length, method, content_type, resource):
    x_headers = 'x-ms-date:' + date
    string_to_hash = method + "\n" + str(content_length) + "\n" + content_type + "\n" + x_headers + "\n" + resource
    bytes_to_hash = bytes(string_to_hash, encoding="utf-8")
    decoded_key = base64.b64decode(shared_key)
    encoded_hash = base64.b64encode(hmac.new(decoded_key, bytes_to_hash, digestmod=hashlib.sha256).digest()).decode()
    authorization = "SharedKey {}:{}".format(customer_id,encoded_hash)
    return authorization

# Build and send a request to the POST API
def post_data(customer_id, shared_key, body, log_type):
    method = 'POST'
    content_type = 'application/json'
    resource = '/api/logs'
    rfc1123date = datetime.datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')
    content_length = len(body)
    signature = build_signature(customer_id, shared_key, rfc1123date, content_length, method, content_type, resource)
    uri = 'https://' + customer_id + '.ods.opinsights.azure.com' + resource + '?api-version=2016-04-01'

    headers = {
        'content-type': content_type,
        'Authorization': signature,
        'Log-Type': log_type,
        'x-ms-date': rfc1123date
    }

    response = requests.post(uri,data=body, headers=headers)
    if (response.status_code >= 200 and response.status_code <= 299):
        print('Accepted')
    else:
        print("Response code: {}".format(response.status_code))

for p in disk_partitions:
    usage = psutil.disk_usage(p.mountpoint)
    system['CounterValue'].append({
        'mountpoint': p.mountpoint,
        'CounterValueFree':usage.free
    })
    body = json.dumps(system)
    print(body)
    post_data(customer_id, shared_key, body, log_type)
    system['CounterValue'] = []
    time.sleep(0.1)


