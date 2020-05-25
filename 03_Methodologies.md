# Methodologies

I utilized a widely adopted approach to performing penetration testing that is effective in testing how well the Offensive Security Exam environments is secured. Below is a breakout of how I was able to identify and exploit the variety of systems and includes all individual vulnerabilities found.

## Information Gathering

The information gathering portion of a penetration test focuses on identifying the scope of the penetration test. During this penetration test, I was tasked with exploiting the exam network. The specific IP addresses were:

Exam Network 

- 192.168.
- 192.168.
- 192.168.
- 192.168.
- 192.168. 
 
## Penetration

The penetration testing portions of the assessment focus heavily on gaining access to a variety of systems. During this penetration test, I was able to successfully gain access to X out of the X systems.

### System IP: 192.168. ()
#### Service Enumeration

The service enumeration portion of a penetration test focuses on gathering information about what services are alive on a system or systems. This is valuable for an attacker as it provides detailed information on potential attack vectors into a system. Understanding what applications are running on the system gives an attacker needed information before performing the actual penetration test.  In some cases, some ports may not be listed.

| Server IP Address | Ports Open |
| ----------------: | ---------: |
| 192.168.          | TCP:       |
| ^                 | UDP:       |

##### Nmap Scan Results:

#### Initial Shell Vulnerability Exploited  

##### Vulnerability Explanation: 

##### Vulnerability Fix: 

##### Severity: 

##### Proof of Concept Code Here: 

##### Local.txt Proof Screenshot:

##### Local.txt Contents:

#### Privilege Escalation

##### Vulnerability Exploited: 

##### Vulnerability Explanation: 

##### Vulnerability Fix: 

##### Severity: 

##### Exploit Code:

##### Proof Screenshot Here:

##### Proof.txt Contents:
 
## Maintaining Access

Maintaining access to a system is important to us as attackers, ensuring that we can get back into a system after it has been exploited is invaluable. The maintaining access phase of the penetration test focuses on ensuring that once the focused attack has occurred (i.e. a buffer overflow), we have administrative access over the system again. Many exploits may only be exploitable once and we may never be able to get back into a system after we have already performed the exploit. 

## House Cleaning

The house cleaning portions of the assessment ensures that remnants of the penetration test are removed. Often fragments of tools or user accounts are left on an organization's computer which can cause security issues down the road. Ensuring that we are meticulous and no remnants of our penetration test are left over is important.

After collecting trophies from the exam network was completed, the student removed all user accounts and passwords as well as the Meterpreter services installed on the system. Offensive Security should not have to remove any user accounts or services from the system.

