

<#
AUTHOR: Tarun Gupta

.SYNOPSIS
 This powershell script fetches poolmembers from all NSX load balancees and send a mail to BIlling team 
    
.DEPENDCY :
    POwreshell Modules required .: IMport-Excel and POwerNSX
#>

<#
Import-Module  'D:\TarunPowercliScripts\ImportExcel-master\ImportExcel-master\ImportExcel.psm1'
Import-MOdule  'D:\TarunPowercliScripts\PowerNSX Module\PowerNSX.psm1'
Add-type -path "$env:D:\TarunPowercliScripts\ImportExcel-master\ImportExcel-master\EPPlus.dll"
#>

# RE and CM NSX Netadmin  Credentails 
$pwstring = ‘Tacdwn!!’
$secpasswd = ConvertTo-SecureString $pwstring -AsPlainText -Force
$NSXCredentials = New-Object System.Management.Automation.PSCredential "netadmin", $secpasswd

#MA NSX Manager Credentails (since netadmin credentails are diffrent for MA NSX Manager)
$pwstring = ‘dO!nc&crq1st’
$secpasswd = ConvertTo-SecureString $pwstring -AsPlainText -Force
$MANSXCredentials = New-Object System.Management.Automation.PSCredential "netadmin", $secpasswd

#$dcvi=@("10.231.232.131","10.229.232.131","10.245.232.131")
$dcnsx=@("10.231.232.141","10.245.232.141","10.229.232.141")
$datacenter=@("CM","RE","MA")


#Get date for file name
$date = get-date -f yyyy-MM-dd

#Making sure A new file is createed eveytime this data is gathered for the same day.

$filepath="C:\script\NSXLoadBalancer_$date.xlsx"
Remove-Item -Path $filepath -Force -ErrorAction SilentlyContinue


for($i=0;$i -lt 3;$i++)
{
Write-host "Fetching data from NSX Manager $($dcnsx[$i])"  -ForegroundColor Green
if($i -eq 2) #if MA NSX MANGER use diffrent credentials
{
Connect-NsxServer $dcnsx[$i] -Credential $MANSXCredentials -DisableVIAutoConnect -WarningAction SilentlyContinue

}
else{
Connect-NsxServer $dcnsx[$i] -Credential $NSXCredentials -DisableVIAutoConnect -WarningAction SilentlyContinue
}



$pools=@()
#Pull down all edge load balancer info
Write-host "Pulling info ... " 
$pools = get-nsxedge | Get-NsxLoadBalancer | Get-NsxLoadBalancerpool 



foreach($pool in $pools)
{
    $poolname = $pool.name
    $poolmembers = $pool.member
    $poolmemberscount=$poolmembers.name.count
    $NSXLBid=$pool.edgeId
    $edgename=Get-NsxEdge -objectId $pool.edgeId

    <#if($poolname.StartsWith("z"))
	{
        $poolname = $poolname.trim("z-")        
    }#>

  $poolcustomer = $edgename.name.substring(0,2)

if ($poolmemberscount -ne 0)
{

foreach($member in $poolmembers)
{
$finaloutput=@()
 $poolMembername =$member.name
 Write-host "working on poolmemeber $($poolMembername)" -ForegroundColor Green
 $poolMemberstate=$member.condition
 $memberscount =$poolMembername.count

 if($member.PSObject.Properties.name -contains "groupingObjectId")
 {
 if($member.groupingObjectId -match "securitygroup")
 {
   $VM= @(Get-NsxSecurityGroup -objectId $member.groupingObjectId |Get-NsxSecurityGroupEffectiveVirtualMachine)
 if($VM.VmName.count -ne 0)
{
foreach ($effectiveVM in $VM.VmName)
{
$output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="securitygroup"
VirtualMachine=$effectiveVM
poolMemberstate=$poolMemberstate
}
$finaloutput +=$output
}
}
else{
$output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="Empty"
VirtualMachine=$null
poolMemberstate=$poolMemberstate
}
$finaloutput +=$output


}
}
 elseif ($member.groupingObjectId -match "vm-")
 {
 $VM=$member.groupingObjectName
 
 $output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="VM"
VirtualMachine=$VM
poolMemberstate=$poolMemberstate
}
$finaloutput +=$output

}
}
elseif($member.PSObject.Properties.name -contains "ipAddress")
{
$ipaddress=$member.ipAddress

$output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="IPAddress"
VirtualMachine=$ipaddress
poolMemberstate=$poolMemberstate
}

$finaloutput +=$output
}
elseif(($member.PSObject.Properties.name -notcontains "ipAddress") -and ($member.PSObject.Properties.name -contains "groupingObjectId"))
{
$output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="EMPTY"
VirtualMachine=$null
poolMemberstate=$poolMemberstate
}
$finaloutput +=$output
}
else
{
$output=[PSCustOmobject][Ordered]@{
Datacenter =$datacenter[$i]
Customer=$poolcustomer
NSXLoadbalancerID=$NSXLBid
Poolname=$poolname
Poolmembercount=$memberscount
poolMemberName= $poolMembername
poolMemberType ="DIFFERENT"
VirtualMachine=$null
poolMemberstate=$poolMemberstate
}
$finaloutput +=$output


}

$finaloutput |Export-Excel -Path $filepath -WorksheetName "$($datacenter[$i])"  -TableName "NSX$($datacenter[$i])" -TableStyle Light9 -Append -AutoSize -AutoFilter




}

    
}


}

}

Write-host " Done......"


