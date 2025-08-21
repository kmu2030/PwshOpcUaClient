<#
The MIT License (MIT)

Copyright(c) 2025 KITA Munemitsu
https://github.com/kmu2030

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the “Software”), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

<#
# About This Script
This script demonstrates PwshOpcUaClient, which connects to an OPC UA server and reads and writes values.

## Usage Environment
Controllers: OMRON Co., Ltd. NX1 (Ver. 1.64 or later), NX5 (Ver. 1.64 or later), NX7 (Ver. 1.35 or later), NJ5 (Ver. 1.63 or later)
IDE:Sysmac Studio 1.62 or later
PowerShell: PowerShell 7.5 or later

## Usage Steps (Simulator)
1.  Run `../setup.ps1`.
    This retrieves the assemblies required by `PwshOpcUaClient` using NuGet.
2.  Open `ExampleReadWriteVariables.smc2` in Sysmac Studio.
3.  Start the simulator and the OPC UA server for simulation.
4.  Generate a certificate on the OPC UA server for simulation.   
    This step is unnecessary if a certificate has already been generated.
5.  Register a user and password for the OPC UA server for simulation.   
    This step is unnecessary if a user has already been registered.
6.  Run `ExampleReadWriteVariables.ps1`.

## Usage Steps (Controller)
1.  Run `../setup.ps1`.
    This retrieves the assemblies required by `PwshOpcUaClient` using NuGet.
2.  Open `ExampleReadWriteVariables.smc2` in Sysmac Studio.
3.  Adjust the project's configuration and settings to match the controller you are using.
4.  Transfer the project to the controller.
5.  Generate a certificate on the controller's OPC UA server.   
    This step is unnecessary if a certificate has already been generated.
6.  Register a user and password for the controller's OPC UA server.   
    This step is unnecessary if a user has already been registered.
7.  Run `ExampleVariableReadWrite.ps1`.
8.  Trust the client certificate on the controller's OPC UA server. Trust the rejected client certificate.   
    This step is unnecessary if you are using anonymous access without signing or encryption for message exchange.
9.  Run `ExampleReadWriteVariables.ps1`.


# このスクリプトについて
このスクリプトはPwshOpcUaClientの例示で、OPC UAサーバに接続して値の読み書きを行う。

## 使用環境
コントローラ : オムロン社製 NX1(Ver.1.64以降), NX5(Ver.1.64以降), NX7(Ver.1.35以降), NJ5(Ver.1.63以降)
IDE         : Sysmac Studio 1.62以降
PowerShell  : PowerShell 7.5以降

## 使用手順 (シミュレータ)
1.  `../setup.ps1`を実行
    `PwshOpcUaClient`が必要とするアセンブリをNuGetで取得。
2.  Sysmac Studioで`ExampleReadWriteVariables.smc2`を開く
3.  シミュレータとシミュレーション用OPC UAサーバを起動
4.  シミュレーション用OPC UAサーバで証明書を生成
    既に生成してある場合は不要。
5.  シミュレーション用OPC UAサーバへユーザーとパスワードを登録
    既に登録してある場合は不要。
6.  `ExampleReadWriteVariables.ps1`を実行

## 使用手順 (コントローラ)
1.  `../setup.ps1`を実行
    `PwshOpcUaClient`が必要とするアセンブリをNuGetで取得。
2.  Sysmac Studioで`ExampleReadWriteVariables.smc2`を開く
3.  プロジェクトの構成と設定を使用するコントローラに合わせる
4.  プロジェクトをコントローラに転送
5.  コントローラのOPC UAサーバで証明書を生成
    既に生成してある場合は不要。
6.  コントローラのOPC UAサーバへユーザーとパスワードを登録
    既に登録してある場合は不要。
7.  `ExampleVariableReadWrite.ps1`を実行
8.  コントローラのOPC UAサーバでクライアント証明書の信頼
    拒否されたクライアント証明書を信頼する。
    Anonymousでメッセージ交換に署名も暗号化も使用しないのであれば不要。
9.  `ExampleReadWriteVariables.ps1`を実行
#>

using namespace Opc.Ua
param(
    [bool]$UserSimulator = $true,
    [string]$ServerUrl = 'opc.tcp://localhost:4840',
    [bool]$UseSecurity = $true,
    [string]$UserName = 'taker',
    [string]$UserPassword = 'chocolatepancakes',
    [double]$Interval = 0.05
)
. "$PSScriptRoot/../PwshOpcUaClient.ps1"

function Main () {
    try {
        $AccessUserIdentity = [string]::IsNullOrEmpty($UserName) `
                                ? (New-Object UserIdentity) `
                                : (New-Object UserIdentity -ArgumentList $UserName, $UserPassword)
        $clientParam = @{
            ServerUrl = $ServerUrl
            UseSecurity = $UseSecurity
            SessionLifeTime = 60000
            AccessUserIdentity = $AccessUserIdentity
        }
        $client = New-PwshOpcUaClient @clientParam

        # The namespace is different between the simulator and the controller.
        $ns = $UserSimulator ? '2' : '4';

        # Define write values.
        $writeValues = New-Object WriteValueCollection
        $writeValue = New-Object WriteValue
        $writeValue.NodeId = New-Object NodeId -ArgumentList "ns=$ns;s=WriteIntVal"
        $writeValue.AttributeId = [Attributes]::Value
        $writeValue.Value = New-Object DataValue
        $writeValues.Add($writeValue)

        # Define read values.
        $readValues = New-Object ReadValueIdCollection
        $readValue = New-Object ReadValueId -Property @{
            AttributeId = [Attributes]::Value
        }
        $readValue.NodeId = New-Object NodeId -ArgumentList "ns=$ns;s=ReadIntVal"
        $readValues.Add($readValue)

        $results = $null
        $diagnosticInfos = $Null
        $exception = $null
        [Int32]$counter = 0
        While ($true) {
            # Write $counter to `WriteIntVal` in the server.
            $_writeValues = $writeValues.Clone()
            $_writeValues[0].Value.Value = $counter
            $results = $null
            $diagnosticInfos = $null
            $response = $client.Session.Write(
                $null,
                $_writeValues,
                [ref]$results,
                [ref]$diagnosticInfos
            )
            if ($null -ne ($exception = ValidateResponse(
                                            $response,
                                            $results,
                                            $diagnosticInfos,
                                            $_writeValues,
                                            'Failed to write.'))
            ) {
                throw $exception
            }

            # Read 'ReadIntVal' from the server.
            $results= New-Object DataValueCollection
            $diagnosticInfos = New-Object DiagnosticInfoCollection
            $response = $client.Session.Read(
                $null,
                [double]0,
                [TimestampsToReturn]::Both,
                $readValues,
                [ref]$results,
                [ref]$diagnosticInfos
            )
            if ($null -ne ($exception = ValidateResponse(
                                            $response,
                                            $results,
                                            $diagnosticInfos,
                                            $readValues,
                                            'Failed to read.'))
            ) {
                throw $exception
            }

            "counter=$counter, ReadIntVal=$($results[0].Value)"
                | Write-Host

            ++$counter
            if ($counter -gt 1000) { break }
            Start-Sleep -Seconds $Interval
        }
    }
    catch {
        $_.Exception
    }
    finally {
        Dispose-PwsOpcUaClient -Client $client
    }
}

class OpcUaFetchException : System.Exception {
    [hashtable]$CallInfo
    OpcUaFetchException([string]$Message,
                        [hashtable]$CallInfo) : base($Message)
    {
        $this.CallInfo = $CallInfo
    }
}

function ValidateResponse {
    param(
        $Response,
        $Result,
        $DiagnosticInfos,
        $Requests,
        $ExceptionMessage
    )

    if (($Results
            | Where-Object { $_ -is [StatusCode]}
            | ForEach-Object { [ServiceResult]::IsNotGood($_) }
        ) -contains $true `
        -or ($Results.Count -ne $Requests.Count)
    ) {
        return [OpcUaFetchException]::new($ExceptionMessage, @{
            Response = $Response
            Results = $Results
            DiagnosticInfos = $DiagnosticInfos
        })
    } else {
        return $null
    }
}

Main
