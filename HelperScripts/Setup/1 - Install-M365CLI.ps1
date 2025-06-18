# Check if Node.JS is installed

try {

    $ErrorActionPreference = "SilentlyContinue"
    $checkNPM = npm list -g
    $ErrorActionPreference = "Continue"

    if ($checkNPM) {

        Write-Host "Node.JS is already installed." -ForegroundColor Green

    }

}
catch {

    # Install Node.JS

    Write-Host "Node.JS is not installed." -ForegroundColor Yellow
    Write-Host "Attempting to install Node.JS..." -ForegroundColor Cyan

    winget install --id=OpenJS.NodeJS  -e

    if ($?) {

        Write-Host "Finished installing Node.JS." -ForegroundColor Cyan

        # Reload path environment variable
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

    }

    else {

        Write-Error -Message "Error installing Node.JS. Please install it manually."

    }

}

$checkNPM = npm list -g

if ($checkNPM -match "@pnp/cli-microsoft365") {

    Write-Host "CLI for Microsoft 365 is already installed." -ForegroundColor Green

}

else {

    Write-Host "CLI for Microsoft 365 is not installed." -ForegroundColor Yellow
    Write-Host "Attempting to install CLI for Microsoft 365..." -ForegroundColor Cyan

    npm i -g @pnp/cli-microsoft365

    if ($?) {

        Write-Host "Finished installing CLI for Microsoft 365." -ForegroundColor Cyan

        # Reload path environment variable
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")

    }

    else {

        Write-Error -Message "Error installing CLI for Microsoft 365. Please install it manually."

    }

}
