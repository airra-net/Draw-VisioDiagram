<#
.SYNOPSIS
    Microsoft Powershell functions for generate Visio Drawing from xml variable parameters.

.DESCRIPTION
    Microsoft Powershell functions for generate Visio Drawing from xml variable parameters.
    Using PSVisio.ps1 script functions and input parameters for create and rendering Visio Drawing.

.PARAMETER VCFasCodeHomeFolder
    Folder Path for VCF as Code scripts and others artefacts.

.PARAMETER DiagramFileName
    Name of Visio Diagram file.

.PARAMETER DiagramParameters
    XML Variable Parameters for create and rendering Visio Drawing.
  
.NOTES
    Version:        0.1
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  28.09.2007
    Purpose/Change: Initial script development
    
...

    Version:        3.2
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  07.09.2022
    Purpose/Change: Begin Reorganize script.

    Version:        3.3
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  02.01.2025
    Purpose/Change: Reorganize script. Add to Line Object Workflow - LinePattern Properties.

    Version:        3.4
    Author:         Andrii Romanenko
    Website:        blogs.airra.net
    Creation Date:  04.01.2025
    Purpose/Change: Reorganize script. Add to Item Object Workflow - LineWeight Properties.
...
   
.EXAMPLE

    Generate Visio Drawing with xml variable and others parameters: 

    .\Draw-VisioDiagram.ps1 -VCFasCodeHomeFolder "D:\VCFasCode" -DiagramFileName "SDDCConceptDiagramV2.vsd" -DiagramParameters $DiagramParameters
#>

Param ( 
    [Parameter(Mandatory)]
    [string]$VCFasCodeHomeFolder,
        
    [Parameter(Mandatory)]
    [string]$DiagramFileName,
    
    [Parameter(Mandatory)]
    [xml]$DiagramParameters        
)

# Step 1.
# Set Location
# Load Script Functions
Set-Location $VCFasCodeHomeFolder
. .\PSVisio.ps1 # Warning! Running scripts must be enabled on your system.

# Step 2.
# Create Visio Application
# Create Document from Blank Template
# Set Active Page
New-VisioApplication | Out-Null
New-VisioDocument | Out-Null
Set-VisioPage | Out-Null

# Step 3.
# Add All Visio Stensils
$StensilsNodes = $DiagramParameters.Diagram.Stensils.ChildNodes
ForEach ($StensilNode in $StensilsNodes) {

    If ($StensilNode.Attributes['AlternativePath'].value -eq 'true') {

        $StensilFilePath = $VCFasCodeHomeFolder + "\" + $StensilNode.Attributes['File'].value
        Add-VisioStensil -Name $StensilNode.Attributes['Name'].value -File $StensilFilePath
    }
    
    Else {
   
        Add-VisioStensil -Name $StensilNode.Attributes['Name'].value -File $StensilNode.Attributes['File'].value
 
    }

}

# Step 4.
# Set Masters Items
$MasterItemsNodes = $DiagramParameters.Diagram.MasterItems.ChildNodes
ForEach ($MasterItemsNode in $MasterItemsNodes) {
 
        Set-VisioStensilMasterItem -Stensil $MasterItemsNode.Attributes['Stensil'].value -Item $MasterItemsNode.Attributes['Item'].value

}

# Step 5.
# Draw items
$ItemsNodes = $DiagramParameters.Diagram.Items.ChildNodes
ForEach ($ItemNode in $ItemsNodes) {

    Switch ($ItemNode.Attributes['Type'].value) {

        "General" { 
                
            Draw-VisioItem -Master $ItemNode.Attributes['Master'].value `
                -X $ItemNode.Attributes['X'].value `
                -Y $ItemNode.Attributes['Y'].value `
                -Width $ItemNode.Attributes['Width'].value `
                -Height $ItemNode.Attributes['Height'].value `
                -FillForegnd $ItemNode.Attributes['FillForegnd'].value `
                -LinePattern $ItemNode.Attributes['LinePattern'].value `
                -Text $ItemNode.Attributes['Text'].value `
                -VerticalAlign $ItemNode.Attributes['VerticalAlign'].value `
                -ParaHorzAlign $ItemNode.Attributes['ParaHorzAlign'].value `
                -CharSize $ItemNode.Attributes['CharSize'].value `
                -CharColor $ItemNode.Attributes['CharColor'].value `
                -LineColor $ItemNode.Attributes['LineColor'].value `
                -LineWeight $ItemNode.Attributes['LineWeight'].value
        }

        "Line" { 
                
            Draw-VisioLine -BeginX $ItemNode.Attributes['BeginX'].value `
                -BeginY $ItemNode.Attributes['BeginY'].value `
                -EndX $ItemNode.Attributes['EndX'].value `
                -EndY $ItemNode.Attributes['EndY'].value `
                -LineWeight $ItemNode.Attributes['LineWeight'].value `
                -LineColor $ItemNode.Attributes['LineColor'].value `
                -LinePattern $ItemNode.Attributes['LinePattern'].value
                
        }
                
        "Text" { 
                
            Draw-VisioText -X $ItemNode.Attributes['X'].value `
                -Y $ItemNode.Attributes['Y'].value `
                -Width $ItemNode.Attributes['Width'].value `
                -Height $ItemNode.Attributes['Height'].value `
                -Text $ItemNode.Attributes['Text'].value `
                -CharSize $ItemNode.Attributes['CharSize'].value `
                -CharStyle $ItemNode.Attributes['CharStyle'].value `
                -CharColor $ItemNode.Attributes['CharColor'].value `
                -LinePattern $ItemNode.Attributes['LinePattern'].value `
                -FillForegndTrans $ItemNode.Attributes['FillForegndTrans'].value 
                
        }
    }

}

# Step 6.
# Resise Page To Fit Contents
Resize-VisioPageToFitContents | Out-Null

# Step 7.
# Save Document
$SaveDiagramFileName = $VCFasCodeHomeFolder + "\" + $DiagramFileName
Save-VisioDocument -File $SaveDiagramFileName | Out-Null

# Step 8.
# Quit Application
Close-VisioApplication | Out-Null
