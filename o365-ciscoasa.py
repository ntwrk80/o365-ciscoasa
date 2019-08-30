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
                print("Test")
                print(ip4s)
                for ip in ip4s:
                    flatIps.extend([(serviceArea, category, ip, tcpPorts, udpPorts)])

        print('IPv4 Firewall IP Address Ranges')
        #print (flatIps)
        currentServiceArea = " "
        groupList = []
        for ip in flatIps:
            serviceArea = ip [0]
            print (f"ServiceArea: {serviceArea} \n")
            if serviceArea != currentServiceArea:
                output.write (asaIpNetworkGroupObject(currentServiceArea,groupList))
                groupList = []
                currentServiceArea = serviceArea
            ipNet = ipaddress.ip_network(ip[2])
            print ("\n")
            groupList += asaIpNetworkObject(ipNet,currentServiceArea)
            output.write (groupList[1] + "\n")
        output.write (asaIpNetworkGroupObject(currentServiceArea,groupList))

def printXML(endpointSets):
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
                flatIps.extend([(serviceArea, category, ip, tcpPorts, udpPorts) for ip in ip4s])
        print('IPv4 Firewall IP Address Ranges')
        #print (flatIps)
        currentServiceArea = " "
        output.write ("<AddressGroups xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns=\"http://tempuri.org/IPAddressGroupsSchema.xsd\">\n")
        for ip in flatIps:
            serviceArea = ip [0]
            if serviceArea != currentServiceArea:
                if currentServiceArea != " ":
                    output.write ("     </AddressGroup>\n")
                currentServiceArea = serviceArea
                output.write (f"     <AddressGroup enabled=\"true\" description=\"Office 365 {serviceArea}\">\n")
            ipNet = ipaddress.ip_network(ip[2])
            ipStart = ipNet[0]
            ipEnd = ipNet[-1]
            output.write (f"          <Range from=\"{ipStart}\" to=\"{ipEnd}\"/>\n")
        output.write ("     </AddressGroup>\n")
        output.write ("</AddressGroups>\n")

def asaIpNetworkGroupObject(groupName,objectList):
    print ("ENTER asaIpNetworkGroupObject\n")
    grpObject = "object-group network " + groupName.lower() + "\n"
    for item in objectList:
        grpObject += f"   network-object object" + item[0] + "\n"
    grpObject += "\n"
    return grpObject

def asaIpNetworkObject(network,productname):
    print ("ENTER asaIPNetworkObject\n")
    ip = str(network.network_address)
    net = str(network.netmask)
    name = "o365." + productname.lower() + "_" + ip
    networkObject = " "
    networkObject = f"  object network {name} \n    subnet {ip} {net} \n   description O365 {productname.lower()} \n\n"
    return name, networkObject

def asaFqdnNetworkObject(fqdn,productname):
    #Started as code from https://www.ifconfig.it
    #fqdn = fqdn.replace ("*","")
    fqdn = re.sub("^\*\.","",fqdn)
    fqdn = re.sub("^.*\*","",fqdn)
    fqdn = re.sub("^\.","",fqdn)
    objname = productname+"_"+fqdn
    print ("object network " + objname)
    print ("fqdn "+fqdn)
    return objname

def slash2sub(ip):
    #Started as code from https://www.ifconfig.it
	sub = ip
	if "/4" in ip:
		sub = ip.replace("/4"," 240.0.0.0")
	elif "/5" in ip:
		sub = ip.replace("/5"," 248.0.0.0")
	elif "/6" in ip:
		sub = ip.replace("/6"," 252.0.0.0")
	elif "/7" in ip:
		sub = ip.replace("/7"," 254.0.0.0")
	elif "/8" in ip:
		sub = ip.replace("/8"," 255.0.0.0")
	elif "/9" in ip:
		sub = ip.replace("/9"," 255.128.0.0")
	elif "/10" in ip:
		sub = ip.replace("/10"," 255.192.0.0")
	elif "/11" in ip:
		sub = ip.replace("/11"," 255.224.0.0")
	elif "/12" in ip:
		sub = ip.replace("/12"," 255.240.0.0")
	elif "/13" in ip:
		sub = ip.replace("/13"," 255.248.0.0")
	elif "/14" in ip:
		sub = ip.replace("/14"," 255.252.0.0")
	elif "/15" in ip:
		sub = ip.replace("/15"," 255.254.0.0")
	elif "/16" in ip:
		sub = ip.replace("/16"," 255.255.0.0")
	elif "/17" in ip:
		sub = ip.replace("/17"," 255.255.128.0")
	elif "/18" in ip:
		sub = ip.replace("/18"," 255.255.192.0")
	elif "/19" in ip:
		sub = ip.replace("/19"," 255.255.224.0")
	elif "/20" in ip:
		sub = ip.replace("/20"," 255.255.240.0")
	elif "/21" in ip:
		sub = ip.replace("/21"," 255.255.248.0")
	elif "/22" in ip:
		sub = ip.replace("/22"," 255.255.252.0")
	elif "/23" in ip:
		sub = ip.replace("/23"," 255.255.254.0")
	elif "/24" in ip:
		sub = ip.replace("/24"," 255.255.255.0")
	elif "/25" in ip:
		sub = ip.replace("/25"," 255.255.255.128")
	elif "/26" in ip:
		sub = ip.replace("/26"," 255.255.255.192")
	elif "/27" in ip:
		sub = ip.replace("/27"," 255.255.255.224")
	elif "/28" in ip:
		sub = ip.replace("/28"," 255.255.255.240")
	elif "/29" in ip:
		sub = ip.replace("/29"," 255.255.255.248")
	elif "/30" in ip:
		sub = ip.replace("/30"," 255.255.255.252")
	elif "/31" in ip:
		sub = ip.replace("/31"," 255.255.255.254",1)
	elif "/32" in ip:
		sub = ip.replace("/32"," 255.255.255.255")
	elif "/1" in ip:
		sub = ip.replace("/1"," 128.0.0.0")
	elif "/2" in ip:
		sub = ip.replace("/2"," 192.0.0.0")
	elif "/3" in ip:
		sub = ip.replace("/3"," 224.0.0.0")
	else:
		sub = ip+" 255.255.255.255"
	return sub


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
        for endpointSet in endpointSets:
            if endpointSet['category'] in ('Optimize', 'Allow'):
                category = endpointSet['category']
                urls = endpointSet['urls'] if 'urls' in endpointSet else []
                tcpPorts = endpointSet['tcpPorts'] if 'tcpPorts' in endpointSet else ''
                udpPorts = endpointSet['udpPorts'] if 'udpPorts' in endpointSet else ''
                flatUrls.extend([(category, url, tcpPorts, udpPorts) for url in urls])
        flatIps = []
        printASA(endpointSets)
    else:
        print('Office 365 worldwide commercial service instance endpoints are up-to-date')

if __name__ == '__main__':
    main(sys.argv)
