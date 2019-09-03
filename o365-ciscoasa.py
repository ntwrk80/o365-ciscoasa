import json
import urllib.request
import uuid
import os
import ipaddress
import argparse
import sys

#Original starting code from Microsoft article:
#(https://docs.microsoft.com/en-us/office365/enterprise/office-365-ip-web-service)
#

# helper to call the webservice and parse the response
def webApiGet(methodName, instanceName, clientRequestId):
    ws = "https://endpoints.office.com"
    requestPath = ws + '/' + methodName + '/' + instanceName + '?clientRequestId=' + clientRequestId
    request = urllib.request.Request(requestPath)
    with urllib.request.urlopen(request) as response:
        return json.loads(response.read().decode())

def printASA(endpointSets):
    flatIps=[]
    with open('O365-CiscoASA-ObjectGroups.txt', 'w') as output:
        for endpointSet in endpointSets:
            if endpointSet['category'] in ('Optimize', 'Allow'):
                ips = endpointSet['ips'] if 'ips' in endpointSet else []
                category = endpointSet['category']
                serviceArea = endpointSet['serviceArea']
                # IPv4 strings have dots while IPv6 strings have colons
                ip4s = [ip for ip in ips if '.' in ip]
                tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
                udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
                #print("Test")
                #print(ip4s)
                for ip in ip4s:
                    flatIps.extend([(serviceArea, category, ip, tcpPorts, udpPorts)])

        print("Converting O365 Endpoints into ASA Groups")
        #print (flatIps)
        currentServiceArea = " "
        groupList = []
        for ip in flatIps:
            serviceArea = ip [0]
            print (f"ServiceArea: {serviceArea} \n")
            if serviceArea != currentServiceArea:
                if currentServiceArea != " ":
                    output.write (asaIpNetworkGroupObject(currentServiceArea,groupList))
                groupList = []
                uniqueIps = []
                currentServiceArea = serviceArea
            if ip[2] not in uniqueIps:
                uniqueIps.append(ip[2])
                ipNet = ipaddress.ip_network(ip[2])
                asaOutput=asaIpNetworkObject(ipNet,currentServiceArea)
                groupList.append(asaOutput)
                output.write (asaOutput[1] + "\n")
        #print("DEBUG: groupList\n")
        #print(groupList)
        #print("DEBGUG \n")
            print (uniqueIps)
            input ("Wait")
        output.write (asaIpNetworkGroupObject(currentServiceArea,groupList))

def asaIpNetworkGroupObject(groupName,objectList):
    #print ("ENTER asaIpNetworkGroupObject\n")
    grpObject = "object-group network " + groupName.lower() + "\n"
    for item in objectList:
        #print ("DEBUG: Print objectList item\n")
        #print (item)
        #print ("DEBUG")
        grpObject += f"   network-object object " + item[0] + "\n"
    grpObject += "\n"
    return grpObject

def asaIpNetworkObject(network,productname):
    #print ("ENTER asaIPNetworkObject\n")
    ip = str(network.network_address)
    net = str(network.netmask)
    name = "o365." + productname.lower() + "_" + ip
    networkObject = " "
    networkObject = f"object network {name} \n    subnet {ip} {net} \n    description O365 {productname.lower()} \n\n"
    returnObject = (name,networkObject)
    return returnObject

def main (argv):
    # Parse
    # path where client ID and latest version number will be stored
    datapath = 'endpoints_clientid_latestversion.txt'
    # fetch client ID and version if data exists; otherwise create new file
    if os.path.exists(datapath):
        with open(datapath, 'r') as fin:
            clientRequestId = fin.readline().strip()
            latestVersion = fin.readline().strip()
    else:
        clientRequestId = str(uuid.uuid4())
        latestVersion = '0000000000'
        with open(datapath, 'w') as fout:
            fout.write(clientRequestId + '\n' + latestVersion)
    version = webApiGet('version', 'Worldwide', clientRequestId)
    if version['latest'] > latestVersion:
        print('New version of Office 365 worldwide commercial service instance endpoints detected')
        # write the new version number to the data file
        with open(datapath, 'w') as fout:
            fout.write(clientRequestId + '\n' + version['latest'])
        # invoke endpoints method to get the new data
        endpointSets = webApiGet('endpoints', 'Worldwide', clientRequestId)
        # filter results for Allow and Optimize endpoints, and transform these into tuples with port and category
        flatUrls = []
        #for endpointSet in endpointSets:
        #    if endpointSet['category'] in ('Optimize', 'Allow'):
        #        category = endpointSet['category']
        #        urls = endpointSet['urls'] if 'urls' in endpointSet else []
        #        tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
        #        udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
        #        flatUrls.extend([(category, url, tcpPorts, udpPorts) for url in urls])
        #flatIps = []
        printASA(endpointSets)
    else:
        print("Office 365 worldwide commercial service instance endpoints are up-to-date.")

if __name__ == '__main__':
    main(sys.argv)
