Personal Proxy Server
---------------------

1. Introduction
   This program is a http proxy server which relay http request from the client and the forward it to the original server/other proxy.

For a standard http request, the client will try to connect to the original server and send the request. And then the server will send the response back to the client.

       request         response
Client -------> Server --------> Client


When you use a proxy server, the client instead of sending the request to the actual server, it'll send the request to the proxy server.
The proxy server will act as a client and try to send the request to the actual server, which will send the response back to the proxy and the proxy then relay it back to the requesting client.

       request        request         response        response
Client -------> Proxy -------> Server --------> Proxy --------> Client


In these scenario we have many benefits, to name a few :
1. The original server only see the proxy ip that connect to them. the client remain anonymous.
2. Multiple client can use a single internet connection.
3. A proxy server could have a caching system so the proxy doen't have to get actual files from the server instead they using the files which resides on their cache, thus improving speed and descreasing bandwith usage. (personal proxy server currently do not implement caching system)


2. Usage
   *. Setting the proxy server
      - Use Proxy Server
        This setting is used if you want to use another proxy server to route the request. if this option is not checked the proxy will connect directly to the server.
      - Proxy Server
        This is the forwarding proxy server ip/address
      - Proxy Port
        This is the forwarding proxy server port
      - Maximum Connection
        Maximum client connection allowed
      - Local Port
        The port which accepting connection for client requests
      - Force Reload
        Force the request to get the latest files
      - Disable Cookie
        Delete cookie from request and response
      - Replace Actual Server
        Replace Server name from response
      - Replace Actual Proxy
        Replace Proxy name from response
      - Replace User Agent
        Replace User Agent from request
      - Authenticate User
        Ask user for password
      - Local Machine Only
        Accept connection from local machine only

   *. Setting the browser
      - Set the browser to use proxy and set the proxy machine ip in the proxy server textbox and the port. use 127.0.0.1 if you using the browser from the same machine as the proxy server.
