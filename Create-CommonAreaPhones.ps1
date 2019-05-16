#Requires -Version 3.0
Param(
   [Parameter(Mandatory=$True)]
   [string]$CsvInputFile,
   [switch]$CustomVerbose
   #[string]$CsvInputFile = "C:\Data\PSHScripts\2N\20140513-113623-603__400000002_1898407_1.399.973.759.978.xml" #Testing purpose
)

#===============================================================================
# Declarations
#===============================================================================
$CSVArray = import-csv $CsvInputFile
[String]$CA_OU = "OU=UC-Devices,OU=Users,OU=BI,DC=eu,DC=boehringer,DC=com" 

#XML Values
[String]$XmlFile = ""
[String]$orderid = "000000001"
[String]$orderpositionid = "0000001"
[String]$activityid = "000000001"
[String]$execution = "MyshopRequestCAPCreation"

#Paths
[String]$XmlOutPath = "D:\Scripts\temp\XR"
[String]$XmlFile = ""

#===============================================================================
# Functions
#===============================================================================
function Get-TimeStamp {    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)    
}

function Generate-SipAddress{
param([string]$name)
    $SipAddress = "sip:$name@boehringer-ingelheim.com"    
    return $SipAddress
}

function Generate-DisplayNumber{
param([string]$number)
    $DisplayNumber =  $number -replace "tel:",""
    return $DisplayNumber
}

Function Return_FailedNode{
<#
    Purpose: Return name for a failed value
#>
param([string]$xmlfile, [string]$CAName)
 
    # Read the xml file  
    [xml]$xmlDoc = (Get-Content $xmlfile)
    [string]$FailedNode = ""

    $xNode = $xmldoc | select-xml -XPath "//uda[uda2[@Value='NOK']]"
    if($xNode){        
        foreach($item in $xnode){
            [string]$NodeValue = $item.node.value
            if($NodeValue -eq $CAName){ return $true }
        }               
    }
    return $false
}

Function Create-XMLResult{
<#
    Purpose: Generate a xml file with the required output for Myshop/UC4 
#>
param( [string]$orderid, [string]$orderpositionid, [string]$activityid, [string]$execution)
 # Set the File Name   
 $fileName = $(get-date -f yyyyMMdd) + "__" + $orderid + "_" + $orderpositionid + "_" + $activityid + ".xml"
 $filePath = $XmlOutPath + "\" + $fileName
 # Create The Document  
 $enc = New-Object System.Text.UTF8Encoding($false)
 $XmlWriter = New-Object System.XMl.XmlTextWriter($filePath,$enc)  
 # Set The Formatting  
 $xmlWriter.Formatting = "Indented"  
 $xmlWriter.Indentation = "4"  
 # Write the XML Decleration  
 $xmlWriter.WriteStartDocument()  
 # Write Root Element  
 $xmlWriter.WriteStartElement("Order")  
      # Write the Document  
      $xmlWriter.WriteElementString("orderid","$orderid") 
      $xmlWriter.WriteElementString("orderpositionid","$orderpositionid") 
      $xmlWriter.WriteElementString("activityid","$activityid")       
      $xmlWriter.WriteElementString("execution","$execution")   
 # Write Close Tag for Root Element  
 $xmlWriter.WriteEndElement() # <-- Closing RootElement  
 # End the XML Document  
 $xmlWriter.WriteEndDocument()  
 # Finish The Document   
 $xmlWriter.Flush()  
 $xmlWriter.Close()
 return $filePath
}


Function Add-XMLNode{
<#
    Purpose: Add a new node to the xml file
#>
param([string]$xmlfile, [string]$Return_CAName, [string]$Return_Status,[string]$Return_StatusDetailed)
 # Read the xml file  

 $fileName = $(get-date -f yyyyMMdd) + "__" + $orderid + "_" + $orderpositionid + "_" + $activityid + "_2.xml"
 $filePath = $XmlOutPath + "\" + $fileName
 If(Test-Path $filePath){
    [xml]$xmlDoc = (Get-Content $filePath)
 }else{
    [xml]$xmlDoc = (Get-Content $xmlfile)
}
 
 $xml_returnvalues = $xmldoc.SelectSingleNode("//Order/returnvalues")
 if(!$xml_returnvalues) { 
    #write "Node not exists"
    $newNodeReturnValues = $xmlDoc.CreateElement("returnvalues")
 }else{
    $newNodeReturnValues = $xml_returnvalues
    #write "Exists"
 }

 $newNodeCAName = $xmlDoc.CreateElement("uda")
 $newNodeCAName.SetAttribute("Name", "StdParam_BND_InterfaceReturn_CAName") 
 $newNodeCAName.SetAttribute("Value", "$Return_CAName") 
 $newNodeReturnValues.AppendChild($newNodeCAName)

 $newNodeStatus = $xmlDoc.CreateElement("uda2")
 $newNodeStatus.SetAttribute(“Name”,”StdParam_BND_InterfaceReturn_Status”)
 $newNodeStatus.SetAttribute(“Value”,”$Return_Status”)
 $newNodeCAName.AppendChild($newNodeStatus)

 $newNodeStatusDetailed = $xmlDoc.CreateElement("uda2")
 $newNodeStatusDetailed.SetAttribute(“Name”,”StdParam_BND_InterfaceReturn_StatusDetailed”)
 $newNodeStatusDetailed.SetAttribute(“Value”,”$Return_StatusDetailed”)
 $newNodeCAName.AppendChild($newNodeStatusDetailed)

 $xmlDoc.LastChild.AppendChild($newNodeReturnValues)

 $xmlDoc.save($filePath)
}

#===============================================================================
# Execution
#===============================================================================

# Step 1
## Creating the CommonAreaPhones Objects


$XmlFile = Create-XMLResult $orderid $orderpositionid $activityid $execution
           
foreach($item in $CSVArray){
    try{ 
        [String]$Return_CAName = $item.CN_name
        [String]$lineuri = $item.CN_Phonenumber
        [String]$RegistrarPool = $item.registrarpool
        [String]$SipAddress = Generate-SipAddress $item.CN_name   
        [String]$DisplayNumber = Generate-DisplayNumber $item.CN_phonenumber
        

        $res = New-CsCommonAreaPhone -LineUri $lineuri -RegistrarPool $RegistrarPool -OU $CA_OU -DisplayName $Return_CAName -DisplayNumber $DisplayNumber -SipAddress $SipAddress
        if($CustomVerbose){Write-Host "$(Get-TimeStamp) Created $($res.DisplayName)"}
    }catch{
        if($CustomVerbose){Write-Host "$(Get-TimeStamp) Error: $_.Exception."}     
        [String]$Return_Status = "NOK"
        [String]$Return_StatusDetailed = $_.Exception
        Add-XMLNode $XmlFile $Return_CAName $Return_Status $Return_StatusDetailed           
    }
}


# Step 2
## Checking CommonAreaPhones replication
if($CustomVerbose){Write-Host "$(Get-TimeStamp) Checking if CA objects have been replicated."}
foreach($item in $CSVArray){
     
    $isFailedCA = Return_FailedNode $XmlFile $item.CN_name

    if($isFailedCA -eq $false){
        if($CustomVerbose){write-host "Waiting for $($item.CN_name)" -nonewline }    
        do{
            $CA = get-cscommonAreaPhone $item.CN_name -ErrorAction 0
            if($CustomVerbose){write-host "." -nonewline}
            sleep -Seconds 15
        }while ($CA -eq $null)
        if($CustomVerbose){write-host " Done" -NoNewline}
        if($CustomVerbose){write-host " "}
        #Updating XML file
        [String]$Return_CAName = $CA.DisplayName
        [String]$Return_Status  = "OK"
        [String]$Return_StatusDetailed = "OK"
        Add-XMLNode $XmlFile $Return_CAName $Return_Status $Return_StatusDetailed |Out-Null
    }else{
        if($CustomVerbose){Write-Host "$(Get-TimeStamp) $($item.Name) marked as failed in XML file"}
    }
}


# Step 3
## Assigning policies to CA
try{ 
    if($CustomVerbose){Write-Host "$(Get-TimeStamp) Assigning policies to each CA object."}   
    foreach($item in $CSVArray){   
        [String]$Return_CAName = $item.CN_name
        [String]$VoicePolicy = $item.CN_VoicePolicy
        [String]$dialplan = $item.CN_dialplan
        [String]$displayname = $item.CN_Name + " " + $item.CN_Displayname                              
        
        #Updating VoicePolicy
        Grant-CsVoicePolicy -Identity $Return_CAName -Policyname $VoicePolicy
        if($CustomVerbose){Write-Host "$(Get-TimeStamp) Granted $VoicePolicy VoicePolicy for $Return_CAName "}
        #Updating DialPlan
        Grant-CsDialPlan -Identity $Return_CAName -Policyname $dialplan
        if($CustomVerbose){Write-Host  "$(Get-TimeStamp) Granted $dialplan DialPlan for $Return_CAName "}
        
        #Updating Displayname
        set-CsCommonAreaPhone -Identity $Return_CAName  -DisplayName $displayname
        if($CustomVerbose){Write-Host "$(Get-TimeStamp) Updated displayname for $Return_CAName to $displayname"}
    }


}catch{
    if($CustomVerbose){Write-Host "$(Get-TimeStamp) CA: $displayname .Error: $_.Exception." }
    $Return_Status = "NOK"
    $Return_StatusDetailed = $_.Exception
    #Create-XMLResult $orderid $orderpositionid $activityid $Return_Status $execution $Return_PersonGID $Return_StatusDetailed $Return_PhoneNumberType $Return_PhoneNumber        
    return "ERROR"    
}

# Step 4
## Assigning PIN to CA
try{ 
    if($CustomVerbose){Write-Host "$(Get-TimeStamp) Checking if CA objects are ready for PIN assignement."}      
    foreach($item in $CSVArray){   
        [String]$Return_CAName = $item.CN_name
        [String]$SipAddress = Generate-SipAddress $item.CN_name

        if($CustomVerbose){write-host "Waiting for $($item.CN_name)" -nonewline }   
        do{
            $res = Get-CsClientPinInfo $SipAddress -EA 0
            if($CustomVerbose){write-host "." -nonewline}
            sleep -Seconds 30
        }while(!$res)
        if($CustomVerbose){write-host " Done" -NoNewline}
        if($CustomVerbose){write-host " "}

        #Assign PIN
        [int]$PIN = Get-Random -Minimum 111111 -Maximum 999999
        $res = Set-CsClientPin -Identity $SipAddress -Pin $PIN
        if($res){
            if($CustomVerbose){Write-Host "$(Get-TimeStamp) Assigned PIN $($res.PIN) for $Return_CAName "}
        }
    }
}catch{
    Write-Host "ERROR"
    if($CustomVerbose){Write-Host "$(Get-TimeStamp) CA: $SipAddress .Error: $_.Exception." }
    $Return_Status = "NOK"
    $Return_StatusDetailed = $_.Exception
    #Create-XMLResult $orderid $orderpositionid $activityid $Return_Status $execution $Return_PersonGID $Return_StatusDetailed $Return_PhoneNumberType $Return_PhoneNumber        
    return "ERROR"    
}
