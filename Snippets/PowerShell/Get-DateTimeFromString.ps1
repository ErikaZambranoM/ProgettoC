# Function that returns a DateTime object from a string
Function Get-DateTimeFromString {
    Param(
        [Parameter(Mandatory = $true)]
        [String]
        $DateTimeString,

        [Parameter(Mandatory = $false)]
        [Switch]
        $ClosestEnd
    )

    Try {
        # Declare the date formats to be supported
        $DateFormats = [String[]]@(
            # Year only
            'yyyy'

            # Month and year
            'M/yyyy',
            'MM/yyyy',

            # Day, month and year
            'd/M',
            'dd/M',
            'd/MM',
            'dd/MM',
            'd/M/yyyy',
            'dd/M/yyyy',
            'd/MM/yyyy',
            'dd/MM/yyyy',

            # Day, month, year and hour
            'd/M HH',
            'dd/M HH',
            'd/MM HH',
            'dd/MM HH',
            'd/M/yyyy HH',
            'dd/M/yyyy HH',
            'd/MM/yyyy HH',
            'dd/MM/yyyy HH',
            'd/M/yyyy H',
            'dd/M/yyyy H',
            'd/MM/yyyy H',
            'dd/MM/yyyy H',

            # Day, month, year, hour and minute
            'd/M H:m',
            'dd/M H:m',
            'dd/MM H:m',
            'd/MM H:m',
            'd/M HH:m',
            'dd/M HH:m',
            'dd/MM HH:m',
            'd/MM HH:m',
            'd/M H:mm',
            'dd/M H:mm',
            'dd/MM H:mm',
            'd/MM H:mm',
            'd/M HH:mm',
            'dd/M HH:mm',
            'dd/MM HH:mm',
            'd/MM HH:mm',
            'd/M/yyyy H:m',
            'dd/M/yyyy H:m',
            'd/MM/yyyy H:m',
            'dd/MM/yyyy H:m',
            'd/M/yyyy HH:m',
            'dd/M/yyyy HH:m',
            'd/MM/yyyy HH:m',
            'dd/MM/yyyy HH:m',
            'd/M/yyyy H:mm',
            'dd/M/yyyy H:mm',
            'd/MM/yyyy H:mm',
            'dd/MM/yyyy H:mm',
            'd/M/yyyy HH:mm',
            'dd/M/yyyy HH:mm',
            'd/MM/yyyy HH:mm',
            'dd/MM/yyyy HH:mm',

            # Day, month, year, hour, minute and second
            'd/M/yyyy HH:m:s',
            'dd/M/yyyy HH:m:s',
            'd/MM/yyyy HH:m:s',
            'dd/MM/yyyy HH:m:s',
            'd/M/yyyy H:m:s',
            'dd/M/yyyy H:m:s',
            'd/MM/yyyy H:m:s',
            'dd/MM/yyyy H:m:s',
            'd/M HH:m:s',
            'dd/M HH:m:s',
            'd/MM HH:m:s',
            'dd/MM HH:m:s',
            'd/M HH:mm:s',
            'dd/M HH:mm:s',
            'd/MM HH:mm:s',
            'dd/MM HH:mm:s',
            'd/M H:mm:s',
            'dd/M H:mm:s',
            'd/MM H:mm:s',
            'dd/MM H:mm:s',
            'dd/MM HH:m:ss',
            'dd/MM HH:mm:ss',
            'dd/MM/yyyy HH:mm:ss'
        )

        # Convert the date string to a DateTime object
        $DateTime = [DateTime]::MinValue
        If ([DateTime]::TryParseExact($DateTimeString, $DateFormats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$DateTime)) {
            $DateTime = $DateTime
        }
        Else {
            Throw "Unable to convert unsupported date format for date string '$DateTimeString'"
        }

        # If the closest end switch is specified, then convert the date to the end of the period
        If ($ClosestEnd) {
            # Determine the matching format
            ForEach ($Format in $DateFormats) {
                If ([DateTime]::TryParseExact($DateTimeString, $Format, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$DateTime)) {
                    $MatchingFormat = $Format
                    Break
                }
            }

            # Add the appropriate amount of time to the date
            Switch ($MatchingFormat) {
                # Year only
                'yyyy' {
                    $DateTime = $DateTime.AddYears(1).AddSeconds(-1)
                    Break
                }

                # Month and year
                { $_ -in ('M/yyyy', 'MM/yyyy') } {
                    $DateTime = $DateTime.AddMonths(1).AddSeconds(-1)
                    Break
                }

                # Day, month and year
                { $_ -in (
                        'd/M',
                        'dd/M',
                        'd/MM',
                        'dd/MM',
                        'd/M/yyyy',
                        'dd/M/yyyy',
                        'd/MM/yyyy',
                        'dd/MM/yyyy'
                    )
                } {
                    $DateTime = $DateTime.AddDays(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year and hour
                { $_ -in (
                        'd/M HH',
                        'dd/M HH',
                        'd/MM HH',
                        'dd/MM HH',
                        'd/M/yyyy HH',
                        'dd/M/yyyy HH',
                        'd/MM/yyyy HH',
                        'dd/MM/yyyy HH',
                        'd/M/yyyy H',
                        'dd/M/yyyy H',
                        'd/MM/yyyy H',
                        'dd/MM/yyyy H'
                    )
                } {
                    $DateTime = $DateTime.AddHours(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year, hour and minute
                { $_ -in (
                        'd/M H:m',
                        'dd/M H:m',
                        'dd/MM H:m',
                        'd/MM H:m',
                        'd/M HH:m',
                        'dd/M HH:m',
                        'dd/MM HH:m',
                        'd/MM HH:m',
                        'd/M H:mm',
                        'dd/M H:mm',
                        'dd/MM H:mm',
                        'd/MM H:mm',
                        'd/M HH:mm',
                        'dd/M HH:mm',
                        'dd/MM HH:mm',
                        'd/MM HH:mm',
                        'd/M/yyyy H:m',
                        'dd/M/yyyy H:m',
                        'd/MM/yyyy H:m',
                        'dd/MM/yyyy H:m',
                        'd/M/yyyy HH:m',
                        'dd/M/yyyy HH:m',
                        'd/MM/yyyy HH:m',
                        'dd/MM/yyyy HH:m',
                        'd/M/yyyy H:mm',
                        'dd/M/yyyy H:mm',
                        'd/MM/yyyy H:mm',
                        'dd/MM/yyyy H:mm',
                        'd/M/yyyy HH:mm',
                        'dd/M/yyyy HH:mm',
                        'd/MM/yyyy HH:mm',
                        'dd/MM/yyyy HH:mm'
                    )
                } {
                    $DateTime = $DateTime.AddMinutes(1).AddSeconds(-1)
                    Break
                }

                # Day, month, year, hour, minute and second
                { $_ -in (
                        'd/M/yyyy HH:m:s',
                        'dd/M/yyyy HH:m:s',
                        'd/MM/yyyy HH:m:s',
                        'dd/MM/yyyy HH:m:s',
                        'd/M/yyyy H:m:s',
                        'dd/M/yyyy H:m:s',
                        'd/MM/yyyy H:m:s',
                        'dd/MM/yyyy H:m:s',
                        'd/M HH:m:s',
                        'dd/M HH:m:s',
                        'd/MM HH:m:s',
                        'dd/MM HH:m:s',
                        'd/M HH:mm:s',
                        'dd/M HH:mm:s',
                        'd/MM HH:mm:s',
                        'dd/MM HH:mm:s',
                        'd/M H:mm:s',
                        'dd/M H:mm:s',
                        'd/MM H:mm:s',
                        'dd/MM H:mm:s',
                        'dd/MM HH:m:ss',
                        'dd/MM HH:mm:ss',
                        'dd/MM/yyyy HH:mm:ss'
                    )
                } {
                    $DateTime = Get-Date $DateTime -Second 59
                    Break
                }

                Default {
                    Throw "Unable to convert unsupported date format for date string '$DateTimeString'"
                }
            }
        }

        Return $DateTime

    }
    Catch {
        Throw
    }
}

Get-DateTimeFromString -DateTimeString '28/09 15:9'
Get-DateTimeFromString -DateTimeString '28/09 15:9' -ClosestEnd