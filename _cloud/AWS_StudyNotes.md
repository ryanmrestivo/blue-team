# Personal Notes
## *Solutions Architect Certification 2021*

[![N|AWS](https://miro.medium.com/max/1248/1*BAsXOt5-QpdWNiKm-lAD-w.png)](https://aws.amazon.com)

## Quick availability reminders.
| Service | Availability |
| ------ | ------ |
| EC2 | 99.99% |
| S3 Standard, S3 Reduced Redundancy| 99.99% |
| S3 Standard-IA, S3 Intelligent Tiering, S3 Glacier Instant Retrieval | 99.9% |
|S3 Glacier Flexible Retrieval and S3 Glacier Deep Archive Class | 99.99%

*When you calculate the availability for multiple products you have to sum the total percentage of individual unavailability. For example for 3EC2 instances.*
` 100% - (10%* 10% * 10%) = 99.9% `

| Availability | Downtime per month |
| ------ | ------ |
| 99% | 7 Hours/month |
| 99.9% | 43 min/month |
| 99.95% | 22 min/month |
| 99.99% | 4.3 min/month -> 52 min/year |
| 99.999% | 30 sec / month |

## Availability for Databases
1. DynamoDB offers 99.99 percent availability and low latency 
2. DynamoDB global tables to achieve even higher availability: 99.999 percent 
3. Multi-AZ RDS -> using Aurora -> Capped to 99.95%

## Some Userful Numbers
* 25 regions before Re:invent 2021.
* 81 Availability Zones before Re:invent 2021.
* 5 VPCs per region (soft-limit)
* 200 Subnets per VPC (soft-limit)
* IPv4 CIDR blocks per VPC is 5 (soft-limit)
* IPv6 CIDR blocks per VPC is 1 (soft-limit)
* 5 Elastic Ip Address per region. (soft-limit)
* VPC must have a mask of at least /16 
* The maximum possible prefix length for an IP subnet is /32
*More information at [AWS Global Infrastructure]*

## For EC2 Instances
* AWS Snowball edge is best suited for IoT applications. There is one for memory optimized and other for computing optimized.
* Amazon Data Lifecycle Manager to make weekly backups of EC2 instances.
* By default, AWS has a limit of 20 instances per region (soft-limit)
* The VM Import/Export tool handles the secure and reliable transfer for a virtual machine between your AWS account and local data center. 
* The instance type determines the size of the instance store available and the type of hardware used for the instance store volumes. 
* Instance store volumes are included as part of the instance's usage cost. You must specify the instance store volumes that you'd like to use when you launch the instance (except for NVMe instance store volumes, which are available by default).

## EC2 Auto Scaling
### Simple Scaling Policies
With a simple scaling policy, whenever the metric rises above the threshold, Auto Scaling simply increases the desired capacity. 
- ChangeInCapacity 
- ExactCapacity 
- PercentChangeInCapacity

After Auto Scaling completes the adjustment, it waits a cool-down period before executing the policy again, even if the alarm is still breaching. The default cool-down period is 300 seconds, but you can set it as high as you want or as low as 0 effectively disabling it. 

### Step Scaling Policies 
If the demand on your application is rapidly increasing, a simple scaling policy may not add enough instances quickly enough. Using a step scaling policy, you can instead add instances based on how much the aggregate metric exceeds the threshold.

With a step scaling policy, you can optionally specify a warm-up time, which is how long Auto Scaling will wait until considering the metrics of newly added instances. The default warm up time is 300 seconds. Note that there are no cool-down periods in step scaling policies. 

### Target Tracking Policies 
Target tracking policies. All you do is select a metric and target value, and Auto Scaling will create a CloudWatch Alarm and a scaling policy to adjust the number of instances to keep the metric near that target. 

### Scheduled Actions 
When you create a scheduled action, you must specify the following: 
- A minimum, maximum, or desired capacity value
- A start date and time
- Only the AMI is defined by the launch configuration. 

## Simple Storage Service (S3)
* All S3 names need to be globally unique.
* You can use prefixes to apply lifecycle rules to only certain objects within a bucket. 
* You should also be aware that there are minimum times (30 days, for instance) an object must remain within one class before it can be moved.
* You can’t transition directly from S3 Standard to Reduced Redundancy. 
* For pre-signed URL the default expiration value is 3,600 seconds (one hour) 
* Glacier guarantees 99.999999999 percent durability 
* Glacier supports archives as large as 40 TB, whereas S3 buckets have no size limit 
* Glacier is encrypted by default
* Glacier IDs are machine-generated IDs. Glacier names don’t have to be unique.
* Glacier vault = S3 bucket.
* In a bucket single object may be no larger than 5 TB. 
* Individual uploads can be no larger than 5 GB. 
* Use Multipart Upload for any object larger than 100 MB. As the name suggests, Multipart Upload breaks a large object into multiple smaller parts and transmits them individually to their S3 target. 
* When using Transfer Acceleration, uploads are routed through geographically nearby AWS edge locations and, from there, routed using Amazon’s internal network. 
* You’re allowed only 100 S3 buckets per account. 
* Security groups are not used with S3 buckets. 
* The principal is the user to be assigned in an attached policy.
* Amazon S3 inventory lets you make audits of the objects in the bucket.
* AWS Backup allows to make backup plans for
    * Aurora
    * Amazon FsX
    * Amazon EFS
    * AWS Storage Gateway
    * EBS
    * RDS
    * DynamoDB.
* S3 Remote acceleration will work for two buckets in different regions.
* RTMP distributions can manage content only from S3 buckets. RTMP is intended for the distribution of video content. 
* Access Analyzer for S3 alerts you to S3 buckets that are configured to allow access to anyone on the internet or other AWS accounts, including AWS accounts outside of your organization.
* Athena lets you perform advanced SQL queries against data stored in S3. 

## Amazon MQ
* Amazon MQ, Amazon SQS, and Amazon SNS are messaging services that are suitable for anyone from startups to enterprises. If you're using messaging with existing applications, and want to move your messaging to the cloud quickly and easily, we recommend you consider Amazon MQ.  
* If you are building brand new applications in the cloud, we recommend you consider Amazon SQS and Amazon SNS.
* Compatible with RabbitMQ

## AWS Lake Formation
AWS Lake Formation is a service that makes it easy to set up a secure data lake in days. A data lake is a centralized, curated, and secured repository that stores all your data, both in its original form and prepared for analysis. 

## Application Load Balanccer
Because an Application Load Balancer terminates incoming TCP connections and creates new connections to your backend targets, it does not preserve client IP addresses all the way to your target code (such as instances, containers, or Lambda code). The source IP address that your targets see in the TCP packet is the IP address of the Application Load Balancer. However, an Application Load Balancer does preserve the original client IP address by removing it from the original packet’s reply address and inserting *it into an HTTP header before it sends the request to your backend over a new TCP connection.*

## Virtual Private Cloud (VPC)
* You can’t change the primary CIDR block after you create your VPC, so think carefully about your address requirements before creating a VPC. 
* You can’t choose your own IPv6 CIDR.
* All IPv6 addresses are reachable from the Internet. 
* The prefix length of an IPv6 VPC CIDR is always /56. 
* VPN connections are always encrypted. 
* You must always assign an IPv4 CIDR block to a subnet, even if you intend to use only IPv6. 
* If your VPC includes an IPv6 CIDR, AWS will automatically add inbound and outbound rules to permit IPv6 traffic. 
* Global Accelerator static addresses are spread across different AWS points-of-presence (POPs) in over 30 countries. These static addresses are also called any-cast addresses because they are simultaneously advertised from multiple POPs. 
* In VPC peering, can’t use it to share Internet gateways or NAT devices. You can, however, use it to share a network load balancer (NLB). 
* In VPC peering If you want to enable peering only between specific subnets, you can specify the subnet CIDR. 
* Inter region VPC peering is not available for some AWS regions. Peering connections between regions have a maximum transmission unit (MTU) of 1,500 bytes and do not support IPv6. 
* AWS Transit Gateway 
    * You can configure a transit gateway to function as a centralized router that controls how traffic is routed among all of your VPCs and on-premises networks. 
    * Isolated VPCs.- When you have multiple VPCs that you want to connect to an on-premises network but want to keep the VPCs isolated from each other. 
    * Isolated VPCs with Shared Services.- If you’re hosting shared services (such as Active Directory or Link Layer Discovery Protocol [LLDP]) in one VPC, you can use a transit gateway to securely share those resources with other networks, while maintaining isolation among them. 
    * Transit Gateway Peering.- You can peer transit gateways together, even between different regions.
    * Multicast.- AWS Transit Gateway supports multicast traffic between VPCs (Note that a multicast receiver can be any EC2 instance, but the sender must be a Nitro instance.)Keep in mind that multicast routes do not appear in any route tables. 
    * Blackhole Routes.- If you want to block a specific route, you can create a blackhole entry for that route in the transit gateway route table. 
* Unlike VPN connections, Direct Connect links are not encrypted. However, because AWS endpoints are all secured by TLS, traffic between your on-premises network and AWS is secure. 
* Elastic Fabric Adapter.- It’s for super fast TCP/IP networking connections for clusters.
* AWS ParallelCluster can automatically manage your Linux-based HPC cluster so that you don’t have to do it manually 
* A subnet cannot span availability zones. 
* Every ENI must have a primary private IP address. 
    * It can have secondary IP addresses, but all addresses must come from the subnet the ENI resides in. 
    * Once created, the ENI cannot be moved to a different subnet. 
    * An ENI can be created independently of an instance and later attached to an instance. 
* Each VPC contains a default security group that can’t be deleted.
* You can create a security group by itself without attaching it to anything. 
* You also attach multiple security groups to the same ENI. 
* An NACL is attached to a subnet, whereas a security group is attached to an ENI. 
* An NACL can be associated with multiple subnets, but a subnet can have only one NACL. 
* An Internet gateway has no management IP address. 
    * It can be associated with only one VPC at a time and so cannot grant Internet access to instances in multiple VPCs. 
    * It is a logical VPC resource and not a virtual or physical router. 
* Every subnet is associated with the main route table by default. 
* A NAT gateway is a VPC resource that scales automatically to accommodate increased bandwidth requirements. A NAT instance can’t do this. 
* A NAT gateway exists in only one availability zone. 
* There are not multiple NAT gateway types. 
* A NAT instance is a regular EC2 instance that comes in different types. 
* The source/destination check on the NAT instance’s ENI must be disabled to allow the instance to receive traffic not destined for its IP and to send traffic using a source address that it doesn’t own. 
* A peering connection can be used only for instance-to-instance communication. You can’t use it to share other VPC resources (except NLB).
* You cannot route through a VPC using transitive routing. 
* The Site-to-Site VPN tool doesn’t use OpenVPN. 

More information about limits at [AWS VPC Limits]

## AWS Artifact
* Is your go-to, central resource for compliance-related information that matters to you. It provides on-demand access to AWS’ security and compliance reports and select online agreements. Reports available in AWS Artifact include our Service Organization Control (SOC) reports, Payment Card Industry (PCI) reports, and certifications from accreditation bodies across geographies and compliance verticals that validate the implementation and operating effectiveness of AWS security controls. Agreements available in AWS Artifact include the Business Associate Addendum (BAA) and the Nondisclosure Agreement (NDA).
* List of services that comply with (PCI DSS) Level 1 for the handling and transmission of credit card data [AWS PCI DSS]

## Systems Manager
* Automation documents let you perform actions against your AWS resources, including taking EBS snapshots. Although called automation documents, you can still manually execute them. 
* A command document performs actions within a Linux or Windows instance.
* A policy document works only with State Manager and can’t take an EBS snapshot.
* There’s no manual document type. 

## For IAM and authentication
* The Action element refers to the kind of action requested (list, create, etc.), the Resource element refers to the particular AWS account resource that’s the target of the policy, and the Effect element refers to the way IAM should react to a request. 
* For IAM policies In general, an explicit deny effect will always overrule an allow. 
* There’s no such thing as an IAM action statement. 
* The top-level command is IAM, and the correct subcommand is get-access-key-last-used. The parameter is identified by --access-last-key-id. 
    * Parameters (not subcommands) are always prefixed with -- characters. 
* X.509 certificates are used for encrypting SOAP requests, not authentication. 
* AWS CloudHSM provides encryption that’s FIPS 140-2 compliant.
* Key Management Service manages encryption infrastructure but isn’t FIPS 140-2 compliant. 
* Security Token Service is used to issue tokens for valid IAM roles
* Secrets Manager handles secrets for third-party services or databases. 
* Identity pools provide temporary access to defined AWS services to your application users 
* IAM policies are global—they’re not restricted to any one region. Policies do, however, require an:
    * action (like create buckets)
    * an effect (allow)
    * and a resource (S3).  
* STS tokens are used as temporary credentials to external identities for resource access to IAM roles.
* Differentiate between Cognito identity pools and user pools.
* Rotation Process:
    * Monitor the use of old keys
    * Deactivate old keys.
    * Delete old keys.
* You attach IAM roles to services in order to give them permissions over resources in other services within your account. 
* IAM keeps five versions of every customer managed policy 
* RunInstances action launches a new instance.
* StartInstances starts an existing instance but doesn’t launch a new one. 
* Granting a user access to use a KMS key to decrypt data requires adding the user to the key policy as a key user. 
* VPC flow logs record source IP address information for traffic coming into your VPC. 
* You can use an IAM policy or SQS access policy to restrict queue access to certain principals or those coming from a specified IP range. 
* A security group to restrict inbound access to authorized sources is sufficient to guard against a UDP-based DDoS attack. 
* You can revoke and rotate both a customer-managed CMK and a customer-provided key at will. You can’t revoke or rotate an AWS-managed CMK or an S3-managed key. 
* AWS-managed CMKs are rotated only once a year. 
* You can install an ACM-generated certificate on a CloudFront distribution or application load balancer. You can’t export the private key of an ACM-generated certificate, so you can’t install it on an EC2 instance. 
* Security Hub checks the configuration of your AWS services against AWS best practices. 
* An IAM group is not a principal. It’s not possible to perform actions under the auspices of a group in the way that you would a role. 
* A role is an IAM principal that doesn’t have a password or access key. 
    * The IAM user that assumes the role can be in the same account as the role or in a different account 
* Users without AWS credentials tend to consume the AWS services that offer resource-based policies. 
* It can take up to 15 minutes between the time an event occurs and when the CloudTrail creates the log file containing the event. 
* AWS Config can alert you when a resource configuration in your AWS account changes. It can also compare your resource configurations against a baseline and alert you when the configuration deviates

## GuardDuty
* Amazon GuardDuty offers threat detection enabling you to continuously monitor and protect your AWS accounts, workloads, and data stored in Amazon Simple Storage Service (Amazon S3).
* Amazon GuardDuty is a regional service.
* Analyzes AWS CloudTrail, VPC Flow Logs, and AWS DNS logs.
* Amazon GuardDuty is priced based on the quantity of AWS CloudTrail Events analyzed and the volume of Amazon VPC Flow Log and DNS Log data analyzed
* Amazon Inspector is an automated security assessment service that helps improve the security and compliance of applications deployed on AWS. Support Linux and Windows Server instances.
* There are no upfront costs and you pay only for the events analyzed, with no additional software to deploy or threat intelligence feed subscriptions required.
* Consider the following actions:
    * The Behavior finding type is triggered by an instance sending abnormally large amounts of data or communicating on a protocol and port that it typically doesn’t. 
    * The Backdoor finding type indicates that an instance has resolved a DNS name associated with a command and control server or is communicating on TCP port 25. 
    * The Stealth finding type is triggered by weakening password policies or modifying a CloudTrail configuration. 
    * The ResourceConsumption finding type is triggered when an IAM user launches an EC2 instance when they’ve never done so. 
* You do not need to stream any logs to CloudWatch Logs for GuardDuty to be able to analyze them. 
* finding is a notification that details the questionable activity 
* Findings can be categorized in:
    * Backdoor.- EC2 compromised with malware, DDoS attacks, TCP port 25
    * Behavior.- Unusual port or large amount of traffic to an external host.
    * Cryptocurrency.- It may be sending, receiving, or mining Bitcoin. 
    * Pentest.- A system is running Kali Linux. Penetration testing
    * Persistence.- IAM user changing resource permissions, security groups, routes or ACLs
    * Policy.- Root credential used or S3 block public access was disabled.
    * Recon.-  This indicates a reconnaissance attack may be underway
    * ResourceConsumption.- Launching an EC2 instance
    * Stealth.- Password policy was weakened.
    * Trojan.- An EC2 may have a trojan installed.
    * UnauthorizedAcces.- Via API or console.
* Notice that every finding type relates to either the inappropriate use of AWS credentials or the presence of malware on an EC2 instance. 
* It’s different from Amazon Macie that helps you discover and protect your sensitive data in Amazon S3 by classifying what data you have and the security, along with the access controls associated with that data.
* Customers can choose to aggregate security findings produced by Amazon GuardDuty across regions by utilizing Amazon CloudWatch Events

## Amazon Inspector
* Amazon Inspector is an agent-based service that looks for vulnerabilities on your EC2 instances. 
* Inspector agent runs an assessment on the instance and analyzes its network, filesystem, and process activity.
* Common Vulnerabilities and Exposures (Linux and Windows) 
* Center for Internet Security Benchmarks  (Linux and Windows) 
* Runtime Behavior Analysis (Linux and Windows) 
* Network Reachability (Linux and Windows)
* Security Best Practices  (Linux Only)
* Types of alerts:
    * High - Fix ASAP
    * Medium - At next possible opportunity
    * Low - Informational

## Amazon Detective
* Detective is designed to help you correlate events and see how a given event affects particular resources 
* Instead of digging though CloudTrail logs manually.

## Security Hub
* Display findings in user-friendly dashboard.
* Security Hub assesses your account against AWS security best practices and the Payment Card Industry Data Security Standard (PCI DSS). 
* Gathers info from Inspector, GuardDuty and Macie.

## Backup and Recovery
* All S3 storage classes except One Zone-Infrequent Access distribute objects across multiple availability zones. 
* To guard against deletion and data corruption, enable S3 versioning 
* To protect your data against multiple availability zone failures—or the failure of an entire region—you can enable cross-region replication between a source bucket in one region and destination bucket in another. 
* One of the easiest ways to back up an EBS volume is to take a snapshot of it. AWS stores EBS snapshots across multiple availability zones in S3. You can either create a snapshot.  
* You can create a Snapshot Lifecycle Policy and specify an interval of 12 or 24 hours. To retain, up to 1,000 snapshots. 
* The Data Lifecycle Manager can automatically create snapshots of an EBS volume every 12 or 24 hours and retain up to 1,000 snapshots. 
* The Standard storage class replicates objects across multiple availability zones in a region, guarding against the failure of an entire zone. 

## Simple Queue Service (SQS)
- A message can be up to 256 KB in size. 
- Can use batching to perform actions on a batch of up to 10 messages at a time. 
- The default visibility timeout is 30 seconds but can range from 0 seconds to 12 hours. 
- The default retention period is 4 days, but you can configure it to be anywhere from 1 minute to 14 days. 
- The default per-queue delay is 0 seconds, but you can set it as long as 15 minutes. 
- Standard Queues -> 120,000 in-flight messages. 
- First-In, First-Out (FIFO) Queues -> 20,000 in-flight messages.  / 3000 per second
- Polling:
    - In short polling, SQS checks only a subset of its servers for waiting messages. 
    - In long polling, SQS ensures that you receive any messages waiting in the queue. it can take up to 20 seconds to return a response. It’s cheaper.
- Dead letter Queues -> Put away a message that no one can process. maxReceiveCount 
- Retention Times:
    - SQS -> 14 days
    - Kinesis Video Stream -> 7 Days
    - Kinesis Data Stream -> 7 Days
    - Kinesis Data Firehose -> 24 hours.
    - You can use a message timer to hide a message for up to 15 minutes 
- Kinesis Data Firehose requires you to specify a destination for a delivery stream. Kinesis Video Streams and Kinesis Data Streams use a producer-consumer model that allows consumers to subscribe to a stream. 
- SQS and Kinesis Data Streams are similar. But SQS is designed to temporarily hold a small message until a single consumer processes it, whereas Kinesis Data Streams is designed to provide durable storage and playback of large data streams to multiple consumers. 
- A single shard gives you writes of up to 1 MB per second, reads 2MB per second

## Cloudtrail
* Cloud Trail always remain activated by creating a service control policy in AWS organizations.
* You can view and search the last 90 days of events recorded by CloudTrail in the CloudTrail console or by using the AWS CLI
* You can download a CSV or JSON file 
* A trail is a configuration that enables delivery of CloudTrail events to an Amazon S3 bucket, CloudWatch Logs, and CloudWatch Events. 
* When you create an organization trail, a trail with the name that you give it will be created in every AWS account that belongs to your organization
    * Users in member accounts will not have sufficient permissions to delete the organization trail, turn logging on or off, change what types of events are logged, or otherwise alter the organization trail in any way.
    * Creating an organization trail helps you define a uniform event logging strategy for your organization. An organization trail is applied automatically to each AWS account in your organization
* A trail enables CloudTrail to deliver log files to your Amazon S3 bucket. By default, when you create a trail in the console, the trail applies to all regions.
* Limit of 5 trails per region. Hard Limit.
* You can configure your trail to send events to CloudWatch Logs. You can then use CloudWatch Logs to monitor your account for specific API calls and events.
* You can enable log encryption and file integrity.
* Can share log with other accounts. Aggregate logs from multiple accounts
* There are three types of events that can be logged in CloudTrail: management events, data events, and CloudTrail Insights events.
* Management Events = Control Plane Operations.
* Data Events = Data plane operations. Data events are often high-volume activities.
* Data events are not logged by default. You have to explicitly request it. Apply additional charges.
* CloudTrail Insights events capture unusual API call rate or error rate activity in your AWS account.
    * Insights events provide relevant information, such as the associated API, error code, incident time, and statistics, that help you understand and act on unusual activity.
* Insights events are disabled by default when you create a trail. You have to explicitly activate it.
* Beginning on April 12, 2019, trails will be viewable only in the AWS Regions where they log events.
* When you use an AWS STS (security token service) region-specific endpoint, the trail in that region delivers only the AWS STS events that occur in that region. For example, if you are using the endpoint sts.us-west-2.amazonaws.com, the trail in us-west-2 delivers only the AWS STS events that originate from us-west-2. 
* Don’t log data events on the bucket that’s storing your CloudTrail logs. Doing so would create an infinite loop! 
* The CloudWatch Agent is different than the legacy CloudWatch Logs agent, which sends only logs and not metrics. AWS recommends using the CloudWatch agent. 
* EventBridge differs from CloudWatch Alarms in that EventBridge takes some action based on specific events, not metric values. 
* AWS Config tracks the configuration state of your AWS resources at a point in time. Think of AWS Config as a time machine. 
    * It can also show you how your resources are related to one another so that you can see how a change in one resource might impact another. 
    * Note that this is different from CloudTrail, which logs events, and from EventBridge, which can alert on events. Only AWS Config gives you a holistic view of your resources and how they were configured at any point in time. 
    * The configuration recorder is the workhorse of AWS Config. It discovers your existing resources, records how they’re configured, monitors for changes, and tracks those changes over time. 
    * Only 1 configurator recorder per region.
    * Configuration recorder -> Configuration item per each resource it monitors.
    * AWS Config maintains configuration items for every resource it tracks, even after the resource is deleted. 
    * You can’t manually delete a configuration item.
* Creating a bucket and subnet are API actions, regardless of whether they’re performed from the web console or AWS CLI. 
* Viewing an S3 bucket and creating a Lambda function are management events, not data events. 
* CloudWatch uses dimensions to uniquely identify metrics with the same name and namespace 
* Basic monitoring sends metrics every five minutes. Detailed monitoring sends them every minute. 
    * This is different from Regular and High resolution!!!
* CloudWatch can store metrics at regular (roundup to 1,5 minute or hour) or high resolution (more accurate) => affect how metrics are delivered.
* Metrics stored at
    * one-hour resolution age out after 15 months. 
    * Five-minute resolutions are stored for 63 days. 
    * One-minute resolution metrics are stored for 15 days. 
    * High resolution metrics are kept for 3 hours. 
* CloudWatch uses a log stream to store log events from a single source. 
    * You can’t set a retention period on a log stream directly.
    * You can’t manually delete log events individually 
    * Every log stream must be in a log group. 
* Log groups store and organize log streams but do not directly store log events. 
* A metric filter extracts metrics from logs but doesn’t store anything. 
* The CloudWatch agent can deliver logs to CloudWatch from a server but doesn’t store logs 
* CloudTrail will not stream events greater than 256 KB in size.
* There’s also a normal delay, typically up to 15 minutes, before an event appears in a CloudWatch log stream. 
* Alarm states:
    * Breaching
    * Missing
    * Ignore
* The recover action migrates the same instance to a new host. 
* The delivery channel must include:
    * S3 Bucket Name
    * Delivery frequency of a Snapshot
    * SNS topic ARN
* You can’t delete the configuration items manually. AWS Config can delete them automatically them after 30 days.
* Metrics() exists. CloudWatch can graph only a time series. 
* When you specify a frequency for periodic checks, you must specify a valid frequency, or else AWS Config will not accept the configuration. 
* Metric filters apply to entire log groups, and you can create a metric filter only after creating the group 

## AWS Control Tower
* AWS Control Tower provides the easiest way to set up and govern a secure, multi-account AWS environment, called a landing zone.
* AWS customers can implement AWS Control Tower, extend governance into new or existing accounts, and gain visibility into their compliance status quickly
* You can automate the setup of your AWS environment with best-practices blueprints for multi-account structure, identity, access management, and account provisioning workflow
* For ongoing governance, you can select and apply pre-packaged policies organization-wide or to specific groups of accounts.
* Guardrails provide governance controls by preventing deployment of resources that don’t conform to selected policies or detecting non-conformance of provisioned resources
* AWS Control Tower automatically implements guardrails using multiple building blocks such as:
    *  AWS CloudFormation to establish a baseline, 
    * AWS Organizations service control policies (SCPs) to prevent configuration changes
    * AWS Config rules to continuously detect non-conformance.
* There is no API everything is managed by the console.
* AWS Landing Zone is an AWS solution offered through AWS Solution Architect, Professional Services, or AWS Partner Network (APN) Partners providing a fully configurable, customer-managed landing zone implementation.
* Organize accounts and implement preventive guardrails using service control policies (SCPs).
* SCPs affect only IAM users and roles that are managed by accounts that are part of the organization.
* An SCP restricts permissions for IAM users and roles in member accounts, including the member account's root user.
* SCPs affect only member accounts in the organization. They have no effect on users or roles in the management account.
* SCPs affect all users and roles in attached accounts, including the root user.
* AWS Security Hub is used by security teams, compliance professionals, and DevOps engineers to continuously monitor and improve the security posture of their AWS accounts and resource

## Amazon Web Application Firewall (WAF)
* Monitors HTTP and HTTPS for ELB and CloudFront.
* AWS WAF is a web application firewall that helps protect your web applications or APIs against common web exploits and bots that may affect availability, compromise security, or consume excessive resources.
* AWS WAF gives you control over how traffic reaches your applications by enabling you to create security rules that control bot traffic and block common attack patterns, such as SQL injection or cross-site scripting.
* AWS WAF, a pre-configured set of rules managed by AWS or AWS Marketplace Sellers to address issues like the OWASP Top 10 security risks and automated bots that consume excess resources, skew metrics, or can cause downtime. 
* You can deploy AWS WAF on 
    * Amazon CloudFront as part of your CDN solution, 
    * The Application Load Balancer that fronts your web servers or origin servers running on EC2, 
    * Amazon API Gateway for your REST APIs
    * AWS AppSync for your GraphQL APIs.
* You can use it for websites not hosted in AWS.
* View history with Cloud Trail
* It supports IPv6
* WAF can block SQL injection attacks against your application, but only if it’s behind an application load balancer. It’s not necessary for the EC2 instances to have an elastic IP address. 
* Both WAF and Shield Advanced can protect against HTTP flood attacks, which are marked by excessive or malformed requests. 
* Shield Advanced includes WAF at no charge. 
* You can use it together with Lambda to:
    * Check the list of known malicious IP
    * Analyze web server logs and identify IP addresses that generate bad or excessive requests.

## AWS Shield
* Helps protect against DDoS attacks
    * Standard.- Layer 3 and 4
    * Advanced.- Included WAF at no charge. Incluye layer 7.
* M;itigates 99 percent of attacks in five minutes or less. 
* It mitigates attacks against CloudFront and Route 53 in less than one second 
* Against Elastic Load Balancing in less than five minutes. 
* Shield usually mitigates all other attacks in less than 20 minutes. 

## Data Encryption
* An AWS-managed CMK automatically rotates once a year. 
    * You can’t disable, rotate, or revoke it. 
    * You can view existing CMKs and create new ones by browsing to the Encryption Keys link in the IAM Dashboard. 
* Encrypting CloudWatch logs requires using AWS CLI.
* When you use your own CMK
    * You can configure key policies to control who may use the key to encrypt and decrypt data. 
    * You can also rotate, disable, and revoke keys. 
* Most AWS services offer encryption of data at rest using KMS-managed keys.
* For S3 you can encrypt:
    * Server-side encryption with S3-managed Keys (SSE-S3) 
    * Server-side encryption with KMS-managed keys (SSE-KMS) 
    * Server-side encryption with customer-provided keys (SSE-C) 
    * Client-side encryption 
* It’s possible to have a bucket containing objects using different encryption options or no encryption at all.
    * Applying default encryption at the bucket level does not automatically encrypt existing objects in the bucket but only those created moving forward. 
* For EBS you can use KMS-managed key.  You cannot encrypt a volume created from unencrypted volume. First create snapshots.
* For EFS you can encrypt:
    * KMS customer master keys to encrypt files 
    * EFS-managed key to encrypt file system metadata (file, directory names)
* For Data in Transit
    * You can use the Amazon Certificate Manager (ACM) to generate a Transport Layer Security (TLS) certificate and then install it on an application load balancer or a CloudFront distribution. 
    * Can’t install certificate directly on an EC2 instance or on-premises server.
    * You can import TLS certificate but use it on us-east-1

## AWS Macie
* Macie is a service that automatically locates and classifies your sensitive data stored in S3 buckets and shows you how it’s being used. Using machine learning.
* Policy findings include changes to a bucket that reduce its security, such as changing a bucket policy or removing encryption. 
* Sensitive data findings classify any sensitive data Macie finds in an S3 bucket. 
* Macie automatically publishes its findings to EventBridge and AWS Security Hub 
* Publishes findings every 15 minutes.

## Cost Optimization Remarks
* You can activate and manage tags from the Cost Allocation Tags page. You can create and apply cost allocation tags to your resources and use them in your budgets as filters. 
* Tags can take up to 24 hours before they appear in the Billing and Cost Management Dashboard. 
* Tags can’t be applied to resources that were launched before the tags were created. 
* You’re allowed two free budgets per account. Subsequent tags will cost around $0.02 each per month. 
* Reports interact with Athena, Redshift, Quiksight. 
* Reports are configured to be regularly written to an S3 bucket 
* AWS Organizations:
    * Share resources AWS Single Sign-on (SSO)
    * Apply IAM rules globally with Service Control Policies (SCP)
    * Create & Mange accounts programmatically.
    * Cloudtrail can be configured to watch events from all the organization..
* Trusted Advisor-> How the account is following the AWS best practices (Upgrade to Business or Enterprise)
    * Checks underutilized or running idle resources
    * Configuration mismatches.
    * Security checks (ex. wide-open S3 bucket)
    * Fault tolerance.
    * Service Limit checks.
* Simple month calculator -> detail calculation. Can be shared with others.
* AWS TOTAL COST of Ownership (TCO calculator) -> Know a cost on premise vs cloud. (3 years)
* Cost Optimization
    * Maximizing Server Density -> Lambda kind of server density.
    * EC2 Reserved Instances. Check the EC2 Reserved Instance Marketplace.
        * Convertible RI -> -54% than on Demand.
        * Standard RI -> 75% than on Demand.
        * Scheduled RI
        * Payment (All Upfront, partial upfront, No Upfront but hourly)
    * Saving Plans-> RI for one to three year period of constant service:
        * Compute Savings plan -> -66% over OnDemand. Any AWS region. EMR, ECS, EKS or Fargate.
        * EC2 Instance Plans -> 72% than onDemand. Use in a single AWS region..
    * EC2 SPOT instances. We define a max price.
        * Spot prices can change without a warning. Instances shut downs.
        * SPOT Instance Interruption -> Terminate, Stop (only with EBS AMI) or Hibernate.
        * SPOT Instance Pool -> unused EC2 instance
        * SPOT Fleet -> Group of Instances (even on-Demand) Can be launched from multiple spot instance pools.
        * Request type:
            * Request (one-time request)
            * Request and Maintain (to maintain a target capacity using a fleet)
            * Reserve for Duration (uninterrupted instance from 1-6hours)
        * Request type 
            * Total target capacity
            * AMI Instance
            * If you want to incorporate Load Balancing
            * If you want to pass user data
        * Define your workload:
            * Define a private AMI
            * Passing user data.
        * SPOT advisor recommends you sample configurations. 
* Remember to down as demand drops -> use average use as an alarm.
* Delete older EBS snapshots -> Use Lifecycle manager.

## Operational Excellence Remarks
* In CloudFormation the logical ID (logical name) must be unique within a template. => After creation is called Physical ID (both case sensitive)
    * A parameter lets you pass custom values instead of hard-coding.
* Passing values between multiple stacks.
    * Nesting Stacks.- Configure templates to create additional stacks.
        * The templates for your nested stacks can contain an Outputs section where you define which values you want to pass back to the parent stack. 
        * Parents can reference this value using the Fn::GetAtt 
    * Exporting Stack Output Values.- Using the Export field in the Output section with Fn::Sub
        * AWS::StackName is a pseudo-parameter that returns the name of the stack. 
        * Can reference in the same account and region with Fn::ImportValue
* You can’t delete a stack if another one is referencing its outputs.
* Stack Update
    * Direct Update-> Upload template. Immediately, modifying only the resources that changed in the template.
    * Change Set-> CloudFormation will display a list of every resource it will add, remove, or modify. 
        * Select to make the changes immediately, delete the change set or do nothing.
        * Useful to compare several different configurations.
* AWS::CloudFormation::Stack => which points to the template used to generate the nested stack.
* Update Behavior
    * Update with no interruption
    * Update with some interruption (retains Physical ID)
    * Replacement (new resource with new PID and arranges dependent resources).
* Stack Policy -> You can’t remove it. Use it to forbid modifications.
    * It has Effect, Action, Principal, Resource and Condtitions.
        * Action:
            * Update:Modify -> Allows updates that don’t modify nor replace the resource.
            * Update:Replace -> Allows updates only if it will replace the resource.
            * Update: Delete -> Allows updates only if it deletes the resource. 
            * Update:* -> Allows all update actions
        * Principal must always be *
        * Resource -> Logical ID preceded by LogicalResourceId/. 
        * Condition -> Resource types
    * A stack policy cannot prevent a principal from updating a resource directly, deleting a resource, or deleting an entire stack -> YOU SHOULD USE IAM POLICIES
* You can temporarily override a stack policy when doing a direct update. 
* CodeCommit benefits:
    * Encryption at rest using KMS and at transit HTTPS or SSH
    * No size limitation for repositories.
    * 2GB limit per file.
    * Integration with other AWS services.
* Codecommit uses IAM policies for access. (Only 2 credentials per user)
    * AWSCodeCommitFullAccess 
    * AWSCodeCommitPowerUser -> Can’t delete repos
    * AWSCodeCommitReadOnly
    * Only IAM principals can interact using git-tools.
* Codecommit doesn’t use resource-based policies.
* You can grant repository access to a role but you can’t assign Git credentials to a role 
* CodeDeploy is a service that can deploy applications to EC2 or on-premises instances. 
* You can use CodeDeploy to deploy binary executables, scripts, web assets, images, and anything else you can store in an S3 bucket or GitHub repository. Also to deploy Lambda functions.
    * Simultaneous deployments
    * Automatic rollback
    * Centralized monitoring
    * Deployment of multiple applications at once.
* CodeDeploy only supports sources S3 or Github.
* A deployment group can be based on an EC2 Auto Scaling group, EC2 instance tags, or on-premises instance tags. -> Create Deployment.
    * In-Place Deployment -> Existing instances. (Mandatory for on-premises)
    * Blue/Green Deployment -> Used to upgrade an existing application with minimal interruption. With Lambda deployments, CodeDeploy deploys a new version of a Lambda function and automatically shifts traffic to the new version. 
        * Requires a classic, application or network load balancer.
* Lambda always use Blue/Green deployments.
* Deployment configuration:
    * OneatAtime. ( succeed 2/3)
    * HalfAtATime. (half of the new)
    * AllAtOnce (at least one)
* Lifecycle Events -> Executed by the agent (Lifecycle event hook)
    * ApplicationStop
    * BeforeInstall
    * AfterInstall 
    * ApplicationStart
    * ValidateService ->Final for deployments without ELB
    * BeforeBlockTraffic
    * AfterBlockTraffic
    * BeforeAllowTraffic
    * AfterAllowTraffic
* Application Specification File (AppSpec) where to copy the files.
    * Version
    * OS
    * Files 
    * Permisions -Only To Linux
* Rollbacks are new deployments. Happen if something fails.
* CodePipeline lets you automate the different stages of your software development and release process (CI/CD)
* Every CodeDeploy pipeline must include at least two stages and can have up to 10. 
* You can have up to 20 actions in the same stage 
* Can use CloudWatch events to detect when a change is made to the repository or bucket. Alternatively, you can have CodePipeline periodically poll for changes. 
* AWS CodeBuild is a managed build service that lets you compile source code and perform unit tests. 
* CodeDeploy doesn’t let you specify a CodeCommit repository as a source for your application files. But you can specify CodeCommit as the provider for the source action and CodeDeploy as the provider for the deploy action. 
* At least one stage must have an action that’s not a source action type. 
* Artifacts are compressed files from the files used during the pipeline stored in an S3 bucket.
* Artifacts can be used as input or output.
* You can use the same bucket for multiple pipelines, but each pipeline can use only one bucket for artifact storage. 
* AWS Systems Manager lets you automatically or manually perform actions against your AWS resources and on-premises servers.
    * Actions.- Manually or Automatically perform actions against AWS resources
        * Automation - Actions against AWS resources.
        * Command - Actions against Linux or Windows Instances
        * Policy - Metadata to collect (software inventory / network config) 
* Session Manager -> To connect to instances without SSH keys or bastions.
* Patch Manager.- Automate patching of Linux and Windows.
* State Manager is a configuration management tool that ensures your instances have the software you want them to have and are configured in the way you define. 
* Insights aggregate health, compliance, and operational details about your AWS resources into a single area of AWS Systems Manager. 
* Trusted advisor alarms that are free for everybody
    * Public access to an S3 bucket, particularly upload and delete access 
    * Security groups with unrestricted access to ports that normally should be restricted, such as TCP port 1433 (MySQL) and 3389 (Remote Desktop Protocol) 
    * Whether you’ve created an IAM user
    * Whether multi-factor authentication is enabled for the root user Public access to an EBS or RDS snapshot 
* The Inventory Manager collects data from your instances, including operating system and application versions. 
* Compliance insights show how the patch and association status of your instances stacks up against the rules you’ve configured. 
* AWS-SetupManagedInstance -> can perform operations on AWS resources and not tasks within an instance. 

## Remarks for Migration
- For VMWARE -> Agentless Discovery Connector Appliance (OVA)
- For Windows/Linux -> Discovery Agent
- AWS Application Discovery Service -> Integrate AWS Migration Hub.
- AWS Control Tower -> Automate Setup, Apply Guardrails for Governance, automate account provisioning, dashboard. 
- AWS Server Migration Service -> AWS SMS for Hyper-V, VMware, Azure VMs. (Virtual Machine ONLY). Agentless based. Long Cutover
- CloudEndure -> Migration and Disaster Recovery, For physical and VMware servers. Cutover to move from staging to production (in minutes). Flexible, Reliable, Highly Automated. Agent based
- Application Migration Service (MGN) -> Integrate CloudEndure and AWS Console.
- AWS Database Migration Service (AWS DMS) -> For databases. To RDS (Re-platform) or Lift and Shift (EC2). Even within RDMs. Even Relational, NoSQL, Analytics and Data Warehouses.
    - Can Merge, Split, Filter Data
    - Support for Snowball. 
- DataSync -> Migration of NFS and SMB data - Fully Managed.
- AWS Transfer - Fully Managed. ->Stores in S3.
- VMWARE CLOUD on AWS -> To use VMware on the cloud.
- VMWARE instances are METAL
- Service Catalog-> Limit the number of services an Administrator authorizes a user. 
    - Products->Portfolio->Assign Products->Add Constraints/ Grant Access (Only created by Administrator)
- To ensure minimal downtime for the database migration, we’re going to use continuous replication of changes (also known as Change Data Capture (CDC)
- AWS Landing Zone begins with four core accounts:
    - AWS Organization Account
    - Shared Services Account 
    - Log Archive Account -> Aggregates CloudTrails and AWS Config
    - Security Account -> GuardDuty, Cross Account and SNS Topics.

## Additional Remarks 
* “Security groups are stateful, while network ACLs are stateless.”
* MariaDB has a page size of 16 KB. To write 200 MB (204,800 KB) of data every second, it would need 12,800 IOPS. Oracle, PostgreSQL, or Microsoft SQL Server, which all use an 8 KB page size, would need 25,600 IOPS to achieve the same throughput. When provisioning IOPS, you must specify IOPS in increments of 1,000.
* General-purpose SSD storage allocates three IOPS per gigabyte, up to 10,000 IOPS. Therefore, to get 600 IOPS, you’d need to allocate 200 GB. Allocating 100 GB would give you only 300 IOPS. The maximum storage size for gp2 storage is 16 TB, so 200 TB is not a valid value. The minimum amount of storage you can allocate depends on the database engine, but it’s no less than 20 GB, so 200 MB is not valid. 
* When you provision IOPS using io1 storage, you must do so in a ratio no greater than 50 IOPS for 1 GB. 
* Multi-AZ deployments using Oracle, PostgreSQL, MariaDB, MySQL, or Microsoft SQL Server replicate data synchronously from the primary to a standby instance. Only a multi-AZ deployment using Aurora uses a cluster volume and replicates data to a specific type of read replica called an Aurora replica. 
* FOR REDSHIFT. The ALL distribution style ensures every compute node has a complete copy of every table. The EVEN distribution style splits tables up evenly across all compute nodes. The KEY distribution style distributes data according to the value in a specified column. There is no distribution style called ODD. 
* The dense storage node type uses fast SSDs, whereas the dense compute node uses slower magnetic storage. 
* FOR REDSHIFT. The dense compute type can store up to 326 TB of data on magnetic storage. The dense storage type can store up to 2 PB of data on solid state drives. 
* FOR DYNAMO DB A single strongly consistent read of an item up to 4 KB consumes one read capacity unit. Hence, reading 11 KB of data per second using strongly consistent reads would consume three read capacity units. 
* FOR DYNAMO DB When you create a table, you can choose to create a global secondary index with a different partition and hash key. A local secondary index can be created after the table is created, but the partition key must be the same as the base table, although the hash key can be different. 
* DynamoDB global tables, which replicate your DynamoDB tables across regions. DynamoDB global tables have an availability of 99.999 percent. 
* Typical database block sizes range from 2 KB to 32 KB. 
* Amazon Redshift uses a block size of 1 MB
* “PostgreSQL is designed for compatibility with Oracle databases.”
* For storage support for DB:
    1. “InnoDB is the only storage engine Amazon recommends for MySQL and MariaDB deployments in RDS and the only engine Aurora supports. 
    2. MyISAM is another storage engine that works with MySQL but it is not compatible with automated backups. 
    3. XtraDB is another storage engine for MariaDB, but Amazon no longer recommends it. 
    4. The PostgreSQL database engine uses its own storage engine by the same name and it is not compatible with other database engines.”
* SPOT Block activate for 6 hours
* SPOT Fleet -> Maintain reliability with SPOT and On-Demand instances.
* “Memory-optimized instances are EBS optimized, providing dedicated bandwidth for EBS storage. Standard instances are not EBS optimized and further tops at 10,000 Mbps disk throughput.”
* Can’t modify a launch configuration YOU CAN modify a launch template.
* Enabling point-in-time recovery gives you an RPO of about five minutes. The recovery time objective (RTO) depends on the amount of data to restore. 
* Geoproximity routing routes users to the location closest to them. Geolocation routing requires you to create records for specific locations or create a default record. A subnet prefix length must be at least /16. 
* Not every CloudFront distribution is optimized for low-latency service. Requests of an edge location will only achieve lower latency after copies of your origin files are already cached. 
* Elastic Container Service is a good platform for micro services.
* Benstalk operations *AREN’T IDEAL FOR MICROSERVICES.*
* RAID optimization is an OS-level configuration and can, therefore, be performed only from within the OS. 
* There is no default node name in a CloudFormation configuration—nor is there a node of any sort. 
* Advance permission from AWS is helpful only for penetration testing operations. 
* Service Catalog helps you audit your resources but doesn’t contribute to ongoing event monitoring. 
* Config is an auditing tool VS CloudTrail tracks API calls. 
* Redis is useful for operations that require persistent session states and or greater flexibility. If you’re after speed, Redis might not be the best choice; in many cases, Memcached will provide faster service. Redis configuration has a rather steep learning curve. 
* Amazon Kendra -> Look for information using natural language
* FSx for Lustre and Elastic File System are primarily designed for access from Linux file systems. 
* Active-Active Architecture:
    1. Use RDS for SQL Server 2016 or 2017 Enterprise Edition. Enable Multi-AZ support and select the Mirroring/Always On option. Select another region for the mirroring option.

## Additional Reference Links:
* [IAM Logic]
* [AWS Labs]
* [Amazon EBS FAQs]
SPL Lab – Introduction to Amazon Elastic Block Store: https://amazon.qwiklabs.com/catalog?keywords=Introduction+to+Amazon+Elastic+Block+Store+(EBS)&format[]=any&level[]=any&duration[]=any&price[]=any&modality[]=any&language[]=any
Amazon EC2 Instance Store: https://docs.aws.amazon.com/AWSEC2/latest/UserGuide/InstanceStorage.html
Load Balancing Documentation: https://aws.amazon.com/documentation/elastic-load-balancing/
Amazon SQS FAQs: https://aws.amazon.com/sqs/faqs/
Elastic Load Balancing FAQs: https://aws.amazon.com/elasticloadbalancing/faqs/
Amazon EC2 FAQs: https://aws.amazon.com/ec2/faqs/
Amazon Route 53 Documentation: https://aws.amazon.com/documentation/route53/
Amazon Glacier Documentation: https://docs.aws.amazon.com/glacier/index.html#lang/en_us
FAQs: Amazon EFS - https://aws.amazon.com/efs/faq/ Amazon S3 - https://aws.amazon.com/s3/faqs/ Amazon Glacier - https://aws.amazon.com/glacier/faqs/ Amazon CloudFront - https://aws.amazon.com/cloudfront/faqs/
Introduction to Amazon S3 Lab – https://amazon.qwiklabs.com/catalog?keywords=Introduction+to+Amazon+Simple+Storage+Service+(S3)&format[]=any&level[]=any&duration[]=any&price[]=any&modality[]=any&language[]=any
Introduction to Amazon EFS – https://amazon.qwiklabs.com/catalog?keywords=Introduction+to+Amazon+Elastic+File+System+(EFS)&format[]=any&level[]=any&duration[]=any&price[]=any&modality[]=any&language[]=any
AWS Well Architected Framework: https://aws.amazon.com/architecture/well-architected/
FAQs: https://aws.amazon.com/elasticloadbalancing/faqs/ https://aws.amazon.com/cloudwatch/faqs/ https://aws.amazon.com/ec2/autoscaling/faqs/ https://aws.amazon.com/autoscaling/faqs/
Introduction to Amazon EC2 Auto Scaling lab: https://amazon.qwiklabs.com/catalog?keywords=Introduction+to+Amazon+EC2+Auto+Scaling&format[]=any&level[]=any&duration[]=any&price[]=any&modality[]=any&language[]=any
Amazon Web Services: Overview of Security Processes: https://d0.awsstatic.com/whitepapers/Security/AWS_Security_Whitepaper.pdf
“AWS Security Best Practices” whitepaper: https://d0.awsstatic.com/whitepapers/aws-security-best-practices.pdf
IAM Best Practices: http://docs.aws.amazon.com/IAM/latest/UserGuide/best-practices.html
Best Practices for Managing AWS Access Keys: http://docs.aws.amazon.com/general/latest/gr/aws-access-keys-best-practices.html
“Protecting Your Data in AWS” video: https://www.youtube.com/watch?v=Tb_W1w_TwLk
“Encrypting Data at Rest” whitepaper: https://d0.awsstatic.com/whitepapers/AWS_Securing_Data_at_Rest_with_Encryption.pdf
“Protecting Data Using Encryption” in Amazon S3 Developer Guide: http://docs.aws.amazon.com/AmazonS3/latest/dev/UsingEncryption.html
“AWS: Overview of Security Processes” whitepaper: https://d0.awsstatic.com/whitepapers/Security/AWS_Security_Whitepaper.pdf
"IAM Policies and Bucket Policies and ACLs! Oh, My! (Controlling Access to S3 Resources) in Amazon Security blog (http://blogs.aws.amazon.com/security/post/TxPOJBY6FE360K/IAM-policies-and-Bucket-Policies-and-ACLs-Oh-My-Controlling-Access-to-S3-Resourc)
“AWS re:Invent 2015: VPC Fundamentals and Connectivity Options” (https://www.youtube.com/watch?v=5_bQ6Dgk6k8)
“Linux Bastion Hosts on the AWS Cloud (https://s3.amazonaws.com/quickstart-reference/linux/bastion/latest/doc/linux-bastion-hosts-on-the-aws-cloud.pdf)
“What is Amazon VPC?” (http://docs.aws.amazon.com/AmazonVPC/latest/UserGuide/VPC_Introduction.html)

More links:
Exam details: https://aws.amazon.com/certification/certified-solutions-architect-associate/
Courses: https://aws.amazon.com/training/course-descriptions/
Whitepapers: https://aws.amazon.com/whitepapers/
Architecture Center: https://aws.amazon.com/architecture/ 
Documentation: https://docs.aws.amazon.com/index.html#lang/en_us

## License

MIT

**Free Software, Hell Yeah!**

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)

   [AWS Global Infrastructure]: <https://aws.amazon.com/about-aws/global-infrastructure/>
   [AWS VPC Limits]: <https://docs.aws.amazon.com/vpc/latest/userguide/amazon-vpc-limits.html>
   [AWS PCI DSS]: <https://aws.amazon.com/compliance/services-in-scope/>
   [IAM Logic]: <https://docs.aws.amazon.com/IAM/latest/UserGuide/reference_policies_evaluation-logic.html#policy-eval-basics>
   [AWS Labs]: <https://wellarchitectedlabs.com/>
   [Amazon EBS FAQs]: <https://aws.amazon.com/ebs/faqs/>
