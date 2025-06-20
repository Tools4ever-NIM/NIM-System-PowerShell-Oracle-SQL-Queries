#
# Oracle SQL Queries.ps1 - IDM System PowerShell Script for Oracle Database
#

$Log_MaskableKeys = @(
    'password'
)


#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'DataProvider'
                type = 'combo'
                label = 'Data Provider'
                description = 'Data Provider for Oracle SQL server'
                table = @{
                    rows = @(
                        # Get ODP.NET_Managed_ODAC122cR1.zip from https://www.oracle.com/database/technologies/odac-downloads.html
                        # Install using elevated command prompt: install_odpm.bat D:\Oracle x64 true
                        @{ id = '0'; display_text = 'Oracle Data Provider for .NET - Managed Driver' }
                    )
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'display_text'
                    }
                }
                value = '0'
            }
            @{
                name = 'ODP_NET_InstallPath'
                type = 'textbox'
                label = 'ODP.NET installation path'
                description = 'Path of ODP.NET installation'
                value = ''
            }
            @{
                name = 'connection_header'
                type = 'text'
                text = 'Connection'
				tooltip = 'Connection information for the database'
            }
            @{
                name = 'DataSource'
                type = 'textbox'
                label = 'Data Source'
                description = 'Data Source of Oracle SQL server'
                value = ''
            }
            @{
                name = 'UseSvcAccountCreds'
                type = 'checkbox'
                label = 'Use credentials of service account'
                value = $true
            }
            @{
                name = 'Username'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                description = 'User account name to access Oracle SQL server'
                value = ''
                hidden = 'UseSvcAccountCreds'
            }
            @{
                name = 'Password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                description = 'User account password to access Oracle SQL server'
                value = ''
                hidden = 'UseSvcAccountCreds'
            }
            @{
                name = 'query_timeout'
                type = 'textbox'
                label = 'Query Timeout'
                description = 'Time it takes for the query to timeout'
                value = '1800'
            }
            @{
                name = 'session_header'
                type = 'text'
                text = 'Session Options'
				tooltip = 'Options for system session'
            }
			@{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                tooltip = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                tooltip = ''
                value = 1
            }
			@{
                name = 'table_header'
                type = 'text'
                text = 'Tables'
				tooltip = 'Query to Table mapping'
            }
			@{
                name = 'table_1_header'
                type = 'text'
                text = 'Table 1'
				tooltip = 'Table 1 Config'
            }
            @{
                name = 'table_1_name'
                type = 'textbox'
                label = 'Table 1 Name'
                description = ''
            }
            @{
                name = 'table_1_query'
                type = 'textbox'
                label = 'Table 1 Query'
                description = ''
				disabled = '!table_1_name'
                hidden = '!table_1_name'
            }
			
			@{
				name = 'table_2_header'
				type = 'text'
				text = 'Table 2'
				tooltip = 'Table 2 Config'
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_name'
				type = 'textbox'
				label = 'Table 2 Name'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_query'
				type = 'textbox'
				label = 'Table 2 Query'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}

			@{
				name = 'table_3_header'
				type = 'text'
				text = 'Table 3'
				tooltip = 'Table 3 Config'
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_name'
				type = 'textbox'
				label = 'Table 3 Name'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_query'
				type = 'textbox'
				label = 'Table 3 Query'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}

			@{
				name = 'table_4_header'
				type = 'text'
				text = 'Table 4'
				tooltip = 'Table 4 Config'
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_name'
				type = 'textbox'
				label = 'Table 4 Name'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_query'
				type = 'textbox'
				label = 'Table 4 Query'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}

			@{
				name = 'table_5_header'
				type = 'text'
				text = 'Table 5'
				tooltip = 'Table 5 Config'
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_name'
				type = 'textbox'
				label = 'Table 5 Name'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_query'
				type = 'textbox'
				label = 'Table 5 Query'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}

			@{
				name = 'table_6_header'
				type = 'text'
				text = 'Table 6'
				tooltip = 'Table 6 Config'
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_name'
				type = 'textbox'
				label = 'Table 6 Name'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_query'
				type = 'textbox'
				label = 'Table 6 Query'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}

			@{
				name = 'table_7_header'
				type = 'text'
				text = 'Table 7'
				tooltip = 'Table 7 Config'
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_name'
				type = 'textbox'
				label = 'Table 7 Name'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_query'
				type = 'textbox'
				label = 'Table 7 Query'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}

			@{
				name = 'table_8_header'
				type = 'text'
				text = 'Table 8'
				tooltip = 'Table 8 Config'
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_name'
				type = 'textbox'
				label = 'Table 8 Name'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_query'
				type = 'textbox'
				label = 'Table 8 Query'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}

			@{
				name = 'table_9_header'
				type = 'text'
				text = 'Table 9'
				tooltip = 'Table 9 Config'
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_name'
				type = 'textbox'
				label = 'Table 9 Name'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_query'
				type = 'textbox'
				label = 'Table 9 Query'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}

			@{
				name = 'table_10_header'
				type = 'text'
				text = 'Table 10'
				tooltip = 'Table 10 Config'
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_name'
				type = 'textbox'
				label = 'Table 10 Name'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_query'
				type = 'textbox'
				label = 'Table 10 Query'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}

			@{
				name = 'table_11_header'
				type = 'text'
				text = 'Table 11'
				tooltip = 'Table 11 Config'
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_name'
				type = 'textbox'
				label = 'Table 11 Name'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_query'
				type = 'textbox'
				label = 'Table 11 Query'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}

			@{
				name = 'table_12_header'
				type = 'text'
				text = 'Table 12'
				tooltip = 'Table 12 Config'
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_name'
				type = 'textbox'
				label = 'Table 12 Name'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_query'
				type = 'textbox'
				label = 'Table 12 Query'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}

			@{
				name = 'table_13_header'
				type = 'text'
				text = 'Table 13'
				tooltip = 'Table 13 Config'
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_name'
				type = 'textbox'
				label = 'Table 13 Name'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_query'
				type = 'textbox'
				label = 'Table 13 Query'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}

			@{
				name = 'table_14_header'
				type = 'text'
				text = 'Table 14'
				tooltip = 'Table 14 Config'
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_name'
				type = 'textbox'
				label = 'Table 14 Name'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_query'
				type = 'textbox'
				label = 'Table 14 Query'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}

			@{
				name = 'table_15_header'
				type = 'text'
				text = 'Table 15'
				tooltip = 'Table 15 Config'
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_name'
				type = 'textbox'
				label = 'Table 15 Name'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_query'
				type = 'textbox'
				label = 'Table 15 Query'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}

			@{
				name = 'table_16_header'
				type = 'text'
				text = 'Table 16'
				tooltip = 'Table 16 Config'
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_name'
				type = 'textbox'
				label = 'Table 16 Name'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_query'
				type = 'textbox'
				label = 'Table 16 Query'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}

			@{
				name = 'table_17_header'
				type = 'text'
				text = 'Table 17'
				tooltip = 'Table 17 Config'
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}
			@{
				name = 'table_17_name'
				type = 'textbox'
				label = 'Table 17 Name'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
				
			}
			@{
				name = 'table_17_query'
				type = 'textbox'
				label = 'Table 17 Query'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}

			@{
				name = 'table_18_header'
				type = 'text'
				text = 'Table 18'
				tooltip = 'Table 18 Config'
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_name'
				type = 'textbox'
				label = 'Table 18 Name'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_query'
				type = 'textbox'
				label = 'Table 18 Query'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}

			@{
				name = 'table_19_header'
				type = 'text'
				text = 'Table 19'
				tooltip = 'Table 19 Config'
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_name'
				type = 'textbox'
				label = 'Table 19 Name'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_query'
				type = 'textbox'
				label = 'Table 19 Query'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}

			@{
				name = 'table_20_header'
				type = 'text'
				text = 'Table 20'
				tooltip = 'Table 20 Config'
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_name'
				type = 'textbox'
				label = 'Table 20 Name'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_query'
				type = 'textbox'
				label = 'Table 20 Query'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
        )
    }

    if ($TestConnection) {
        Open-OracleSqlConnection $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log verbose "Done"
}


function Idm-OnUnload {
    Close-OracleSqlConnection
}

#
# CRUD functions
#

$ColumnsInfoCache = @{}

$SqlInfoCache = @{}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $connection_params = ConvertFrom-Json2 $SystemParams

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # Get all tables and views in database
            #

            Open-OracleSqlConnection $SystemParams
            
            #
            # Output list of supported operations per table/view (named Class)
            #
            for ($i = 0; $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name".length -gt 0)
                {
                    @(
                        [ordered]@{
                            Class = $connection_params."table_$($i)_name"
                            Operation = 'Read'
                            'Source type' = 'Query'
                            'Primary key' = ''
                            'Supported operations' = 'R'
                            'Query' = $connection_params."table_$($i)_query"
                        }
                    )
                    }
            }
        }
        else {
            # Purposely no-operation.
        }

    }
    else {

        if ($GetMeta) {
            #
            # Get meta data
            #

            Open-OracleSqlConnection $SystemParams

            @()

        }
        else {
            #
            # Execute function
            #

            Open-OracleSqlConnection $SystemParams

            for ($i = 0; $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name" -eq $class)
                {
                    $class_query = ($connection_params."table_$($i)_query" -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join ' '
                    break
                }
            }
            
            $sql_command = New-OracleSqlCommand $class_query
            Invoke-OracleSqlCommand $sql_command
            Dispose-OracleSqlCommand $sql_command

        }

    }

    Log verbose "Done"
}


#
# Helper functions
#

function New-OracleSqlCommand {
    param (
        [string] $CommandText
    )

    $sql_command = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($CommandText, $Global:OracleSqlConnection)
    $sql_command.CommandTimeout = $connection_params.query_timeout
    $sql_command.BindByName = $true

    return $sql_command
}


function Dispose-OracleSqlCommand {
    param (
        [Oracle.ManagedDataAccess.Client.OracleCommand] $SqlCommand
    )

    $SqlCommand.Dispose()
}

function Invoke-OracleSqlCommand {
    param (
        [Oracle.ManagedDataAccess.Client.OracleCommand] $SqlCommand,
        [string] $DeParamCommand
    )

    # Streaming
    function Invoke-OracleSqlCommand-ExecuteReader {
        param (
            [Oracle.ManagedDataAccess.Client.OracleCommand] $SqlCommand
        )
        
        Log verbose "Executing Query: $($SqlCommand.CommandText)"
        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            $hash_table = [ordered]@{}

            foreach ($column_name in $column_names) {
                $hash_table[$column_name] = ""
            }

            $obj = New-Object -TypeName PSObject -Property $hash_table

            # Read data
			if($data_reader.HasRows) {
				while ($data_reader.Read()) {
					foreach ($column_name in $column_names) {
						$obj.$column_name = if ($data_reader[$column_name] -is [System.DBNull]) { $null } else { $data_reader[$column_name] }
					}

					# Output data
					$obj
				}
			} else { [PSCustomObject]$hash_table }
        }

        $data_reader.Close()
    }

    # Non-streaming (data stored in $data_set)
    try {
        Invoke-OracleSqlCommand-ExecuteReader $SqlCommand
    }
    catch {
        Log error "Query Failure: $_"
        throw $_
    }

    Log debug "Done"
}


function Open-OracleSqlConnection {
    param (
        [string] $ConnectionParams
    )

    $connection_params = ConvertFrom-Json2 $ConnectionParams

    Add-Type -Path "$($connection_params.ODP_NET_InstallPath -replace '[/\\]?odp.net[/\\]?$')\odp.net\managed\common\Oracle.ManagedDataAccess.dll"

    $cs_builder = New-Object Oracle.ManagedDataAccess.Client.OracleConnectionStringBuilder

    # Use connection related parameters only
    $cs_builder['Data Source'] = $connection_params.DataSource

    if ($connection_params.UseSvcAccountCreds) {
        # None
    }
    else {
        $cs_builder['User ID']  = $connection_params.Username
        $cs_builder['Password'] = $connection_params.Password
    }

    $connection_string = $cs_builder.ConnectionString

    if ($Global:OracleSqlConnection -and $connection_string -ne $Global:OracleSqlConnectionString) {
        Log verbose "OracleSqlConnection connection parameters changed"
        Close-OracleSqlConnection
    }

    if ($Global:OracleSqlConnection -and $Global:OracleSqlConnection.State -ne 'Open') {
        Log warn "OracleSqlConnection State is '$($Global:OracleSqlConnection.State)'"
        Close-OracleSqlConnection
    }

    if ($Global:OracleSqlConnection) {
        #Log debug "Reusing OracleSqlConnection"
    }
    else {
        Log verbose "Opening OracleSqlConnection '$connection_string'"

        try {
            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connection_string)
            $connection.Open()

            $Global:OracleSqlConnection       = $connection
            $Global:OracleSqlConnectionString = $connection_string

            $Global:ColumnsInfoCache = @{}
            $Global:SqlInfoCache = @{}
        }
        catch {
            Log error "Connection Failure: $($_)"
            throw $_
        }

        Log verbose "Done"
    }
}


function Close-OracleSqlConnection {
    if ($Global:OracleSqlConnection) {
        Log verbose "Closing OracleSqlConnection"

        try {
            $Global:OracleSqlConnection.Close()
            $Global:OracleSqlConnection = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log verbose "Done"
    }
}
