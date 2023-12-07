# [Deprecated]ContainerInsightsLogCollector

## Introduction
This script is designed to collect related logs from OMS Agent pod enabled with Container Insights.
https://docs.microsoft.com/en-us/azure/azure-monitor/containers/container-insights-overview

## Prerequistes
Bash environment with kubectl connected to target cluster.

## Usage
`wget https://raw.githubusercontent.com/ripenedcat/azure_monitor/main/ContainerInsightsLogCollector/ContianerInsightsLogCollector.sh && bash ContianerInsightsLogCollector.sh`

## Function
1. Collect:
- ama-logs/omsagent deployment
- configmap 
2. For ama-logs-rs-xxxxx or omsagent-rs-xxxxx pod , collect:
- describe pod
- container log
- agent log
- cim log
3. For 1 random ama-logs-xxxxx or omsagent-xxxxx pod, collect:
- describe pod
- container omsagent log
- agent log
- cim log
