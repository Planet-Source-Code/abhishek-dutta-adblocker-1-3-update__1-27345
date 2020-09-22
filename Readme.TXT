Before using this program please back up your hosts file. Please check the hosts.svr file(in data directory) for names of servers you might not want to block.

Please do not add too many server names in the blocked list, because an inflated hosts file may actually slow down your internet connection. (around 300 server names are already there in the blocked list, but try to keep it below 500 ). Also, please don't add too many server names and their ip addresses in the DNS Cache because that would again get added up in the hosts file. Only add ips & site names of sites you visit frequently. The format should be -
ip address <space> sitename eg:- 216.5.163.42 www.planetsourcecode.com

If you want to enble selective cookie deletion, then right click and tray icon and go to settings. Click on the cookies tab and add the name of the server from which no cookie is to be deleted. For eg: You don't want to delete any cookies from Planetsourcecode.com, so what you do is - add planetsourcecode to the list, and uncheck "delete all cookies" option.

Reducing the timer interval can cause unpredictable behaviour, atleast in terms of blocking popups. The popup blocked count show in the systray will almost always be wrong and valid popups will also get closed. Besides, it will also be a wastage of system resouces. Please experiment with your system settings to get optimum output. It would be safer to disable popup blocking while browsing sites where you would rather want to see popups (eg: image galleries etc.)

All files in Data & Image are required for proper working of the program and it should be present in the application directory.

Please refer to Readme.WRI for details.

Please report bugs, if you find any.Thanks for downloading - Enjoy.

/data/cookies.svr - names of sites from which cookies are not to be deleted
/data/hosts.svr - names of blocked sites
/data/hosts.bak - backup of original hosts file
/data/popups.svr - captions of popups not to be closed
/data/dns.svr - ip ads and server names in DNS Cache
/images/*.gif - web server image files
/images/*.ico - icons used in this project
/adservers.txt - list of over 600 ad servers
