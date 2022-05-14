# XLL_Phishing

## Introduction
With Microsoft's [recent announcement](https://docs.microsoft.com/en-us/deployoffice/security/internet-macros-blocked) regarding the blocking of macros in documents originating from the internet (email AND web download), attackers have began aggressively exploring other options to achieve user driven access (UDA). There are several considerations to be weighed and balanced when looking for a viable phishing for access method:

1. Complexity - The more steps that are required on the user's part, the less likely we are to be successful. 
2. Specificity - Are most victim machines susceptible to your attack? Is your attack architecture specific? Does certain software need to be installed?
3. Delivery - Are there network/policy mitigations in place on the target network that limit how you could deliver your maldoc? 
4. Defenses - Is application whitelisting enforced?
5. Detection - What kind of AV/EDR is the client running?

These are the major questions, however there are certainly more.  Things get more complex as you realize that these factors compound each other; for example, if a client has a web proxy that prohibits the download of executables or DLL's, you may need to stick your payload inside a container (zip, iso, etc).  Doing so can present further issues down the road when it comes to detection.  More robust defenses require more complex combinations of techniques to defeat.

## What are XLL's?
XLL's are DLL's, specifically crafted for Microsoft Excel. XLL's provide a very attractive option for UDA given that they are executed by Microsoft Excel, a very commonly encountered software in client networks; as an additional bonus, because they are executed by Excel, our payload will almost assuredly bypass Application Whitelisting rules because a trusted application (Excel) is executing it.  The downside of course is that there are very few legitimate uses for XLL's, so it SHOULD be a very easy box to check for organizations to block the download of that file extension through both email and web download. Sadly many organzations are years behind the curve and as such XLL's stand to be a viable method of phishing for some time.

There are a series of different events that can be used to execute code within an XLL, the most notable of which is xlAutoOpen.  The full list may be seen [here](https://docs.microsoft.com/en-us/office/client-developer/excel/add-in-manager-and-xll-interface-functions): 

![image](https://user-images.githubusercontent.com/91164728/168409483-8433b74d-cb40-4b67-bbc1-6ed1e55f5a9c.png)

Upon double clicking an XLL, the user is greeted by this screen:

![image](https://user-images.githubusercontent.com/91164728/168409898-528faad9-2801-44ba-8128-0107db7dd6a3.png)

This single dialog box is all that stands between the user and code exection; with fairly thin social engineering, code execution is all but assured.

## Delivery
Delivery of a payload is a serious consideration in context of UDA.  There are two primary methods we will focus on:

1. Email
2. Web Delivery

### Email
Either via attaching a file or including a link to a website where a file may be downloaded, email is a critial part of the UDA process. Over the years many organizations (and email providers) have matured and enforced rules to protect users and organizations from malicious attachments.  Mileage will vary, but organizations now have the capability to:

1. Block executable attachments (exe, dll, xll, MZ headers overall)
2. Block containers like ISO/IMG which are mountable and may contain executable content
3. Examine zip files and block those containing exectuable content
4. Block zip files that are password protected
5. More

Fuzzing an organizations email rules can be an important part of an engagement, however care must always be taken so as to not tip one's hand that a Red Team operation is ongoing and information is actively being gathered.

For the purposes of this article, it will be assumed that the organization has robust email attachment rules that prevent the delivery of an XLL payload via attachment.  We will pivot and look at Web Delivery.

### Web delivery
Email will still be used in this attack vector, however rather than sending an attachment it will be used to send a link to a website to the user.  Web proxies and network mitigations regarding allowed file download types can differ from those enforced in regards to attached files.  For the purposes of this article, it is assumed that the target prevents the download of executable files (MZ headers).  Luckily for us, while the client's web proxy prevents the download of executable files, it allows zip files containing executables.  These zip files are actively scanned for malware; to include prompting the user for the password for password-protected zip files, however there is not a blanket-deny rule for exectuables as there is for 
