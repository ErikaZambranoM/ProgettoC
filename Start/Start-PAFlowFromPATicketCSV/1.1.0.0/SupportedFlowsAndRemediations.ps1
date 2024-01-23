<# SupportedFlowsAndRemediations
    Array of flows supported by this script for resubmission. Each flow must be a PSCustomObject and have the following properties:
        - DisplayName (Mandatory):
            DisplayName of the flow
        - ResubmitIfEmptyFlowErrors (Mandatory):
            Boolean value indicating if the Flow should be resubmitted by default if Flow Error Details are not present.
        - ResubmitIfUnsupportedFlowErrors (Mandatory):
            Boolean value indicating if the Flow should be resubmitted by default if Flow Error Details are present but not supported by remediation actions.
        - AMSIdentifier (Optional):
            Identifier of the object from managed applications as named inside ticket description (Transmittal Number, Document Number). It must be a string.
            If not set, the property won't get extracted from ticket description and won't be added to the output.
        - PreventiveCheckAction (Optional)
            ScriptBlock to run before resubmitting the Flow to check if it should be resubmitted or not.
            If the ScriptBlock returns $False, the Flow will not be resubmitted.
        - Remediation (Optional):
            PSCustomObject containing details about how and when to run a specific function to solve the issue
            or perform preliminary actions to resubmit the flow.
            If not specified, the flow will be resubmitted without performing any action.
            It must be a PSCustomObject with the following properties:
                - Name (Mandatory):
                    Name of the remediation action to perform. It must be a string.
                - ErrorToRemediate (Mandatory):
                    Error to remediate. It must be a string that can be found in the error details of the flow.
                    If set to '*', the remediation action will be performed for any error.
                - ExecutionOrder (Mandatory):
                    Integer value indicating the order in which the remediation action must be performed.
                    The lower the value, the earlier the remediation action will be performed.
                - IsSPOConnectionRequired (Mandatory):
                    Boolean value indicating if the remediation action requires a connection to the SharePoint Online site provided in PATicket description.
                - Action (Mandatory):
                    Action to perform to remediate the error. It must be a scriptblock.
                    Output of the ScriptBlock to be passed to main script must be provided in a PSCustomObject with the following properties:
                        - Output (Mandatory):
                            Output of the remediation action.
                        - FlowResubmissionResponse (Optional):
                            Response of the flow resubmission.
                        - FlowNeedResubmission:
                            Boolean value indicating if the Flow still needs automatic resubmission after remediation action.
                            If not specified, the flow will be resubmitted by default.
    ! IMPORTANT !
        If a remediation action finds (based on logics in provided ScriptBlock) that a Flow run should not be resubmitted, the returned value must end with 'No need to resubmit the flow.'
#>
# Convert to xml
# bring here functions
# [System.Collections.ArrayList]

@(
    # VDM - Site Independent - Document Disciplines Task Creation
    [PSCustomObject][Ordered]@{
        DisplayName                     = '*Document Disciplines Task Creation'
        ResubmitIfEmptyFlowErrors       = $true
        ResubmitIfUnsupportedFlowErrors = $false
        AMSIdentifier                   = 'Document Number'
        PreventiveCheckAction           = {
            # Nested progress bar
            $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
            $CurrentProgressBarId = $ParentProgressBarId + 1
            $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
            $PercentComplete = (8 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Running PreventiveCheckAction')
                Status          = ('Searching Document')
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting

            # Get Document from PATicket description and check it on Vendor Document List and Process Flow Status List.
            $Global:VDM_Document = Get-VDMDocument -FullDocumentNumber $($PATicket.AMSIdentifier.AMSIdentifierValue) -SPOConnection $SPOConnection
            $PercentComplete = (75 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Running PreventiveCheckAction')
                Status          = ('Checking Document Status')
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting

            # If Document is not in 'Start Commenting' or 'Commenting' status, there is no need to resubmit the Flow.
            If ($Global:VDM_Document.PFSL_Item.Status -ne 'Start Commenting' -and $Global:VDM_Document.PFSL_Item.Status -ne 'Commenting')
            {
                $PreventiveCheckOutput = ("Document '{0}' (Index {1}) is in '{2}' status. No need to resubmit the Flow." -f
                    $Global:VDM_Document.PFSL_Item.TCM_DN,
                    $Global:VDM_Document.PFSL_Item.Index,
                    $Global:VDM_Document.PFSL_Item.Status
                )
                Write-Host $PreventiveCheckOutput -ForegroundColor Yellow
                $ProgressBarStatus = 'Resubission not needed'
                $OutputToReturn = $false
            }
            Else
            {
                Write-Host 'Flow resubmission needed.' -ForegroundColor DarkBlue
                $ProgressBarStatus = 'Resubission needed'
                $OutputToReturn = $true
            }

            $PercentComplete = (100 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Returning PreventiveCheckAction Output')
                Status          = $ProgressBarStatus
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting
            Return $OutputToReturn
        }
        Remediations                    = @(
            # Remove User from Disciplines'
            [PSCustomObject][Ordered]@{
                Name                    = 'Remove User from Disciplines'
                ErrorToRemediate        = 'Referenced User or Group (*) is not found.'
                ExecutionOrder          = 1
                IsSPOConnectionRequired = $true
                Action                  = {
                    # Nested progress bar
                    $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
                    $CurrentProgressBarId = $ParentProgressBarId + 1
                    $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
                    $PercentComplete = (8 / 100 * 100)
                    $ProgressBarSplatting = @{
                        Activity        = ('Running PreventiveCheckAction')
                        Status          = ('Searching Users...')
                        PercentComplete = $([Math]::Round($PercentComplete))
                        Id              = $CurrentProgressBarId
                        ParentId        = $ParentProgressBarId
                    }
                    Write-Progress @ProgressBarSplatting

                    # Remove user from Disciplines list
                    $UserRemovalOutput = Remove-PAErrorUserFromDisciplines -ErrorMessages $PATicketFlowErrorDetails.FlowErrorDetails.ErrorMessages -SPOConnection $SPOConnection

                    $PercentComplete = (100 / 100 * 100)
                    $ProgressBarSplatting = @{
                        Activity        = ('Running PreventiveCheckAction')
                        Status          = ('Returning output...')
                        PercentComplete = $([Math]::Round($PercentComplete))
                        Id              = $CurrentProgressBarId
                        ParentId        = $ParentProgressBarId
                    }
                    Write-Progress @ProgressBarSplatting
                    # Object output to return
                    [PSCustomObject][Ordered]@{
                        RemediationActionOutput = $UserRemovalOutput
                    }
                }
            },
            # Too many tasks assigned to user
            [PSCustomObject][Ordered]@{
                Name                    = 'Too many tasks assigned to user'
                ErrorToRemediate        = 'The request exceeded allowed limits.'
                ExecutionOrder          = 2
                IsSPOConnectionRequired = $false
                Action                  = {
                    # Warn about user having too many tasks assigned. Run cleaning script before resubmitting this flow.
                    Write-Host ('Maximum tasks assigned to user. Run cleaning script before resubmitting this flow.' -f $PATicket.PATicketID, $PATicket.FlowDisplayName, "`n") -ForegroundColor Yellow
                    $TotalSkippedTicket++
                    # Object output to return
                    [PSCustomObject][Ordered]@{
                        # Temporary set to always resubmit the flow (when a single user is assigned too many tasks because in too many Disciplines)
                        FlowNeedResubmission    = $true # Default: $true
                        RemediationActionOutput = 'Automatic resubmission needed.' #Default: 'Automatic resubmission needed.' #Otherwise: 'SKIPPED'
                    }
                }
            },
            # Invalid DueDate
            [PSCustomObject][Ordered]@{
                Name                    = 'Invalid DueDate'
                ErrorToRemediate        = "Schema validation has failed. Validation for field 'DueDate', on entity 'Task' has failed: DueDate cannot be earlier than the StartDate"
                ExecutionOrder          = 3
                IsSPOConnectionRequired = $true
                Action                  = {
                    # Splatting parameters for New-JSONBodyForHTTPFlowRemediation
                    $RemediationParameters = @{
                        FlowDisplayName               = '*Document Disciplines Task Creation'
                        RemediationName               = 'Invalid DueDate'
                        SupportedFlowsAndRemediations = $SupportedFlowsAndRemediations
                        PATicket                      = $PATicket
                        SPOConnection                 = $SPOConnection
                        SpecificRemediationArguments  = $Global:VDM_Document
                    }

                    # Get body for HTTP flow
                    $TaskCreationFlowTriggerBody = New-JSONBodyForHTTPFlowRemediation @RemediationParameters

                    # Trigger specified flow with specified body if body is not empty
                    If ($TaskCreationFlowTriggerBody.StartsWith('{') -and $TaskCreationFlowTriggerBody.EndsWith('}'))
                    {
                        # Trigger specified flow with specified body
                        $FlowResubmissionResponse = Invoke-PAHTTPFlow -Uri $AMSResubmitFlowUri -Body $TaskCreationFlowTriggerBody -Method 'POST'

                        <# Create a dummy response if response is empty
                        If (-not $FlowResubmissionResponse)
                        {
                            $FlowResubmissionResponse = [PSCustomObject]@{
                                ResubmitActionStatusCode = $null
                                ResubmitActionResult     = $null
                                ResubmissionType         = 'Manual'
                                LinkToResubmittedRun     = $null
                                FlowResubmissionDateTime = $null
                                LinkToRun                = $null
                                FlowRunDateTime          = $null
                            }
                        }
                        #>

                        # Object output to return
                        [PSCustomObject][Ordered]@{
                            FlowResubmissionResponse = $FlowResubmissionResponse
                            FlowNeedResubmission     = $false
                            RemediationActionOutput  = 'Flow manually triggered by remediation. No need to resubmit the Flow.'
                        }
                    }
                    Else
                    {
                        # Object output to return
                        [PSCustomObject][Ordered]@{
                            FlowNeedResubmission    = $false
                            RemediationActionOutput = $TaskCreationFlowTriggerBody
                        }
                    }
                }
            },
            # Missing DueDate
            [PSCustomObject][Ordered]@{
                Name                    = 'Missing DueDate'
                ErrorToRemediate        = "the value provided for date time string '' was not valid"
                ExecutionOrder          = 4
                IsSPOConnectionRequired = $true
                Action                  = {
                    # Splatting parameters for New-JSONBodyForHTTPFlowRemediation
                    $RemediationParameters = @{
                        FlowDisplayName               = '*Document Disciplines Task Creation'
                        RemediationName               = 'Invalid DueDate'
                        SupportedFlowsAndRemediations = $SupportedFlowsAndRemediations
                        PATicket                      = $PATicket
                        SPOConnection                 = $SPOConnection
                        SpecificRemediationArguments  = $Global:VDM_Document
                    }

                    # Get body for HTTP flow
                    $TaskCreationFlowTriggerBody = New-JSONBodyForHTTPFlowRemediation @RemediationParameters

                    # Trigger specified flow with specified body if body is not empty
                    If ($TaskCreationFlowTriggerBody.StartsWith('{') -and $TaskCreationFlowTriggerBody.EndsWith('}'))
                    {
                        # Trigger specified flow with specified body
                        $FlowResubmissionResponse = Invoke-PAHTTPFlow -Uri $AMSResubmitFlowUri -Body $TaskCreationFlowTriggerBody -Method 'POST'

                        <# Create a dummy response if response is empty
                        If (-not $FlowResubmissionResponse)
                        {
                            $FlowResubmissionResponse = [PSCustomObject]@{
                                ResubmitActionStatusCode = $null
                                ResubmitActionResult     = $null
                                ResubmissionType         = 'Manual'
                                LinkToResubmittedRun     = $null
                                FlowResubmissionDateTime = $null
                                LinkToRun                = $null
                                FlowRunDateTime          = $null
                            }
                        }
                        #>

                        # Object output to return
                        [PSCustomObject][Ordered]@{
                            FlowResubmissionResponse = $FlowResubmissionResponse
                            FlowNeedResubmission     = $false
                            RemediationActionOutput  = 'Flow manually triggered by remediation. No need to resubmit the Flow.'
                        }
                    }
                    Else
                    {
                        # Object output to return
                        [PSCustomObject][Ordered]@{
                            FlowNeedResubmission    = $false
                            RemediationActionOutput = $TaskCreationFlowTriggerBody
                        }
                    }
                }
            },
            # Throttling
            [PSCustomObject][Ordered]@{
                Name                    = 'Throttling'
                ErrorToRemediate        = 'Too many requests.'
                ExecutionOrder          = 5
                IsSPOConnectionRequired = $false
                Action                  = {
                    # Object output to return
                    [PSCustomObject][Ordered]@{
                        RemediationActionOutput = 'Automatic resubmission needed.'
                        FlowNeedResubmission    = $true
                    }
                }
            }
        )
    },

    # VDL Changes Propagation
    [PSCustomObject][Ordered]@{
        DisplayName                     = 'VDL Changes Propagation'
        ResubmitIfEmptyFlowErrors       = $true
        ResubmitIfUnsupportedFlowErrors = $false
    },

    <# VDL Delete Record
        Template Site Flows: K484, 4346
        Updated Site Flows:
            A2201
            K461
            K439
            4245
            4245
            4305
            43X4
            4274
            4285
            43P4
            4355
            4191
            K475
            43U4
            4325
            K482
    #>
    [PSCustomObject][Ordered]@{
        DisplayName                     = 'VDM * - Delete Record'
        ResubmitIfEmptyFlowErrors       = $true
        ResubmitIfUnsupportedFlowErrors = $true
        AMSIdentifier                   = 'Folder Name'
        PreventiveCheckAction           = {
            # Nested progress bar
            $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
            $CurrentProgressBarId = $ParentProgressBarId + 1
            $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
            $PercentComplete = (8 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Running PreventiveCheckAction')
                Status          = ('Searching Document...')
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting

            # Get Document from PATicket description and check it on Vendor Document List and Process Flow Status List.
            $Global:VDM_Document = Get-VDMDocument -FullDocumentNumber $($PATicket.AMSIdentifier.AMSIdentifierValue) -SPOConnection $SPOConnection -CSVListExpirationMinutes 5 -ErrorAction SilentlyContinue #-ErrorVariable Global:CommandError
            $PercentComplete = (75 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Running PreventiveCheckAction')
                Status          = ('Checking Document...')
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting

            # If Document is present in 'Vendor Document List', it has already been recreated and there is no need to resubmit the Flow.
            If (!(-not $Global:VDM_Document.VDL_Item))
            {
                $PreventiveCheckOutput = ("Document '{0}' (Index {1}) has been recreated in 'Vendor Document List'. No need to continue." -f
                    $Global:VDM_Document.VDL_Item.TCM_DN,
                    $Global:VDM_Document.VDL_Item.Index
                )
                Write-Host $PreventiveCheckOutput -ForegroundColor DarkBlue
                $ProgressBarStatus = 'Resubission not needed'
                $OutputToReturn = $false
            }

            If ($OutputToReturn -ne $false)
            {
                If (!(-not $Global:VDM_Document.PFSL_Item))
                {
                    # If Document is not in 'Placeholder' status, there is no need to resubmit the Flow.
                    Switch ($Global:VDM_Document.PFSL_Item.Status)
                    {
                        'Placeholder'
                        {
                            $PreventiveCheckOutput = 'Document deletion must be completed.'
                            Write-Host $PreventiveCheckOutput -ForegroundColor DarkBlue
                            $ProgressBarStatus = 'Resubission needed'
                            $OutputToReturn = $true
                        }

                        'Deleted'
                        {
                            $PreventiveCheckOutput = ("Document is already in 'Deleted' status. Checking leftovers to be removed...")
                            Write-Host $PreventiveCheckOutput -ForegroundColor Yellow
                            $ProgressBarStatus = 'Resubission needed'
                            $OutputToReturn = $true
                        }

                        Default
                        {
                            $PreventiveCheckOutput = ("Document '{0}' (Index {1}) is in '{2}' status. Can't proceed with deletion." -f
                                $Global:VDM_Document.PFSL_Item.TCM_DN,
                                $Global:VDM_Document.PFSL_Item.Index,
                                $Global:VDM_Document.PFSL_Item.Status
                            )
                            Write-Host $PreventiveCheckOutput -ForegroundColor Yellow
                            $ProgressBarStatus = 'Resubission not needed'
                            $OutputToReturn = $false
                        }

                    }
                }
                Else
                {
                    $PreventiveCheckOutput = ("Document not found in 'Process Flow Status List'. Can't proceed with deletion but it should be already completed successfully." -f
                        $Global:VDM_Document.PFSL_Item.TCM_DN,
                        $Global:VDM_Document.PFSL_Item.Index
                    )
                    Write-Host $PreventiveCheckOutput -ForegroundColor Yellow
                    $ProgressBarStatus = 'Resubission not needed'
                    $OutputToReturn = $false
                }
            }

            $PercentComplete = (100 / 100 * 100)
            $ProgressBarSplatting = @{
                Activity        = ('Returning PreventiveCheckAction Output')
                Status          = $ProgressBarStatus
                PercentComplete = $([Math]::Round($PercentComplete))
                Id              = $CurrentProgressBarId
                ParentId        = $ParentProgressBarId
            }
            Write-Progress @ProgressBarSplatting
            Return $OutputToReturn
        }
        Remediations                    = @(
            # Generic Error
            [PSCustomObject][Ordered]@{
                Name                    = 'Generic Error'
                ErrorToRemediate        = '*'
                ExecutionOrder          = 1
                IsSPOConnectionRequired = $true
                Action                  = {
                    # Nested progress bar
                    $ParentProgressBarId = ([Array]$Global:ProgressBarsIds.Keys)[0]
                    $CurrentProgressBarId = $ParentProgressBarId + 1
                    $Global:ProgressBarsIds[$CurrentProgressBarId] = $true
                    $PercentComplete = (9 / 100 * 100)
                    $ProgressBarSplatting = @{
                        Activity        = ('Running Remediation Action')
                        Status          = ('Deleting Document...')
                        PercentComplete = $([Math]::Round($PercentComplete))
                        Id              = $CurrentProgressBarId
                        ParentId        = $ParentProgressBarId
                    }
                    Write-Progress @ProgressBarSplatting

                    # Remove VDM Document
                    $DeleteVDMDocument_Output = Remove-VDMDocument -VDMDocument $Global:VDM_Document -SPOConnection $SPOConnection
                    $PercentComplete = (100 / 100 * 100)
                    $ProgressBarSplatting = @{
                        Activity        = ('Running Remediation Action')
                        Status          = ('Returning output...')
                        PercentComplete = $([Math]::Round($PercentComplete))
                        Id              = $CurrentProgressBarId
                        ParentId        = $ParentProgressBarId
                    }
                    Write-Progress @ProgressBarSplatting
                    # Object output to return
                    [PSCustomObject][Ordered]@{
                        FlowNeedResubmission    = $false
                        RemediationActionOutput = $DeleteVDMDocument_Output ?? 'Success'
                    }
                }
            }
        )

    }
    <#
    # DD/VDM - Send Transmittal
    [PSCustomObject][Ordered]@{
        DisplayName                     = '*Send Transmittal*'
        ResubmitIfEmptyFlowErrors       = $false
        ResubmitIfUnsupportedFlowErrors = $false
        AMSIdentifier                   = 'Transmittal Number'
        PreventiveCheckAction           = {

            <#
            # Verificare ogni attributo mandatorio e coerenza dati tra lista e file

            # Filter TransmittalQueue_Registry for PA.TrnNumber
            $TransmittalRegistryItem = TransmittalQueue_Registry | Where {$_.Title -eq PA.TrnNumber}
                If ($TransmittalRegistryItem.Count != 1) {Return error}
                Se ($TransmittalRegistryItem.TransmittalStatus = Sent)
                    {Tk ok, non fare nulla}
                Else
                    //Filter TransmittalQueueDetails_Registry for PA.TrnNumber
                    //    ForEach Document
                    //        Filter Main List for Document
                    //            if Document.lastTranmittalDate > QueueRegistry.lastTranmittalDate
                    # Filter TransmittalQueueDetails_Registry for PA.TrnNumber
                    $Counter = 0
                    $TransmittalRegistryDetailsItems = TransmittalQueueDetails_Registry | Where {$_.TransmittalID -eq PA.TrnNumber}
                    ForEach ($Document in $TransmittalRegistryDetailsItems)
                        # Check DetailsRegistry for duplicated Documents
                        $TransmittalRegistryDetailsDuplicatedDocuments = TransmittalQueueDetails_Registry | Where {$_.TCM_DN -eq Document.TCM_DN -and $_.Rev -eq $Document.Rev}
                        If ($TransmittalRegistryDetailsDuplicatedDocuments.Count -gt 1)
                        {
                             #Check each Document on main list to see if all have more recent date
                            $CDL_Document = $CDL | Where {$_.TCM_DN -eq $Document.TCM_DN -and $_.Rev -eq $Document.Rev}
                            If ($CDL_Document.lastTranmittalDate = $null)
                                {Analizzare; Break}
                            Else
                                {$Counter++}
                        }
                        Else
                        {Analizzare}
                    }
            If ($Counter = $TransmittalRegistryDetailsItems.Count){'Tk ok, non fare nulla'}
###end multirow comment

            Return $false
        }
    }
    #>

    # VDM - Site Independent - Document Process
    #[PSCustomObject][Ordered]@{
    #    DisplayName                     = 'VDM - Site Independent - Document Process'
    #    ResubmitIfEmptyFlowErrors       = $false
    #    ResubmitIfUnsupportedFlowErrors = $false
    #}
)