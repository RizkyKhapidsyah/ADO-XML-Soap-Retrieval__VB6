ADO, XML Data Retrieval and SQL execute thru' Soap Service
===========================================================

Requirement of the SOAP server
------------------------------
1. Make sure the server is pre-installed with windows 2000 profesional/server operating system. Windows 98 suxx because the dll compilation made the dll buggy with type library stuffs.

2. The server had to have PWS/IIS 4.0 and above

3. Please compile the MyDll first using VB. You don't have to register the dll because it will be registered after compilation. But if it still not registered then you can go to command prompt, change directory to the directory containing the dll and type>
- Regsvr32 MyDll.dll
Put /u to unregister the dll.

4. After compile, use SOAP toolkit to generate the WSDL file. In Listener URL type the web address where the wsdl will reside. eg: http://localhost/MyWebCom/SoapCom/

5. Create web directory where the WSDL/wsml file resides. Right Click to MyWebCom Directory and go to Web Sharing Tab. 
- Set Access = Read
- Application permission = Execute

6. Run the web server. eg: Open the Administrative Tools> Services. And Start the "World Wide Web Publishing Service".

7. Remember, during compilation of the dll, you can NEVER overwrite the existing dll while the WEB SERVER IS RUNNING. To overwrite the dll, stop the web service first. (i.e. stop the World Wide Web Publishing Service service)


Requirement of the SOAP Client
------------------------------
1. The Client has to register the MSSOAP dll. (file name = MSSOAP1.dll). register the dll with regsvr32. (You are encouraged to install the dll in the folder C:\Program Files\Common Files\MSSoap\Binaries\)

2. Run the client and test the connection. (make sure the Server is running!)

3. The client can execute one SQL statement at a time.



Ok, happy coding!!

Sincerely,
Ahmad Abdul Rahman
(matdp2014@yahoo.com)