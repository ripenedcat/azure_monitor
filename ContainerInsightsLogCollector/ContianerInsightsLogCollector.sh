#! /bin/bash

datevalue=$(date +"%y%m%d_%H%M%S")
path=$datevalue'_containerinsightslog'

old_agent_name="omsagent"
new_agent_name="ama-logs"
current_agent_name=""

list_oms_pods() {
    pods_output=$(kubectl get pods -n kube-system)
    if grep -q "$old_agent_name" <<<"$pods_output"; then
        current_agent_name=$old_agent_name
    elif grep -q "$new_agent_name" <<<"$pods_output"; then
        current_agent_name=$new_agent_name
    else
        echo "${pods_output}"
        echo "==========="
        echo >&2 "No $old_agent_name or $new_agent_name pods found, Aborted. Please use the script when Azure Container Insights is onboarded"
        exit 1
    fi
    echo "current Container Insights pods status:"
    kubectl get pods -n kube-system | grep $current_agent_name
}

log_collector() {
    echo "start log collecting"
    mkdir $path
    echo "<log collection> $current_agent_name deployment "
    kubectl get deployment $current_agent_name-rs -n kube-system -o yaml >$path/deployment.txt 2>&1
    echo "<log collection> application map "
    kubectl get configmaps container-azm-ms-agentconfig -o yaml -n kube-system >$path/configmap.txt 2>&1

    RSOMSPOD=$(kubectl get pods -n kube-system | grep "$current_agent_name"-rs | awk '{print $1}')

    #rs-pod collection
    mkdir $path/$RSOMSPOD
    echo "<log collection> $RSOMSPOD container log "
    kubectl logs $RSOMSPOD -n kube-system --timestamps >$path/$RSOMSPOD-containerlog.log 2>&1
    echo "<log collection> described $RSOMSPOD"
    kubectl describe pod $RSOMSPOD -n kube-system >$path/$RSOMSPOD-describe.log 2>&1

    echo "<log collection> $RSOMSPOD mdsd log "
    kubectl cp -n kube-system $RSOMSPOD:/var/opt/microsoft/linuxmonagent/log $path/$RSOMSPOD 1>/dev/null
    echo "<log collection> $RSOMSPOD cim log "
    kubectl cp -n kube-system $RSOMSPOD:/var/opt/microsoft/docker-cimprov/log $path/$RSOMSPOD 1>/dev/null

    for OMSPOD in $(kubectl get pods -n kube-system | grep "$current_agent_name" | awk '{print $1}'); do
        if [ $OMSPOD == $RSOMSPOD ]; then
            continue
        else
            mkdir $path/$OMSPOD
            echo "*********************************" >>$path/$OMSPOD-containerlog.log
            kubectl logs $OMSPOD -c $current_agent_name -n kube-system --timestamps >$path/$OMSPOD-containerlog.log 2>&1
            echo "<log collection> $OMSPOD container log "
            kubectl logs $OMSPOD -c $current_agent_name-prometheus -n kube-system --timestamps >>$path/$OMSPOD-containerlog.log 2>&1

            echo "<log collection> described $OMSPOD"
            kubectl describe pod $OMSPOD -n kube-system >$path/$OMSPOD-describe.log 2>&1

            echo "<log collection> $OMSPOD mdsd log "
            kubectl cp -n kube-system $OMSPOD:/var/opt/microsoft/linuxmonagent/log $path/$OMSPOD 1>/dev/null
            echo "<log collection> $OMSPOD cim log "
            kubectl cp -n kube-system $OMSPOD:/var/opt/microsoft/docker-cimprov/log $path/$OMSPOD 1>/dev/null

            break
        fi
    done

    #tar -zcvf $path.tar.gz $path/ 1>/dev/null
    #rm -rf $path/

    echo "log collection complete at $path"
}

pack_logs() {
    if [ $(command -v zip | wc -l) == "1" ]; then
        zip -q -r $path.zip $path/ 1>/dev/null
        echo "packed at $path.zip"
    else
        tar -zcvf $path.tar.gz $path/ 1>/dev/null
        echo "log packed at $path.tar.gz"
    fi
}

echo "***********************************************"
echo "Welcome to use Container Insgihts log collector"
echo "***********************************************"

#check kubectl installed
command -v kubectl >/dev/null 2>&1 || {
    echo >&2 "kubectl is not installed, exiting"
    exit 1
}

#get current cluster
current_cluster=$(kubectl config view | grep current)
echo -n "Current cluster is "
echo ${current_cluster/current-context:/}

#list oms pods
list_oms_pods
log_collector
pack_logs
echo "Script finished."
