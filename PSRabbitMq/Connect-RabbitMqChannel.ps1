Function Connect-RabbitMqChannel {
    <#
    .SYNOPSIS
        Create a RabbitMQ channel and bind it to a queue

    .DESCRIPTION
        Create a RabbitMQ channel and bind it to a queue

    .PARAMETER Connection
        RabbitMq Connection to create channel on

    .PARAMETER Exchange
        Optional PSCredential to connect to RabbitMq with

    .PARAMETER ExchangeType
        Specify the Exchange Type to be Explicitly declared as non-durable, non-autodelete, without any option.
        Should you want more specific Exchange, create it prior connecting to the channel, and do not specify this parameter.

    .PARAMETER Key
        Routing Keys to look for

        If you specify a QueueName and no Key, we use the QueueName as the key

    .PARAMETER QueueName
        If specified, bind to this queue.

        If not specified, create a temporal queue

    .PARAMETER Durable
        If queuename is specified, this needs to match whether it is durable

    .PARAMETER Exclusive
        If queuename is specified, this needs to match whether it is Exclusive

    .PARAMETER AutoDelete
        If queuename is specified, this needs to match whether it is AutoDelete

    .EXAMPLE
        $Channel = Connect-RabbitMqChannel -Connection $Connection -Exchange MyExchange -Key MyQueue
 #>
    [outputType([RabbitMQ.Client.Framing.Impl.Model])]
    [cmdletbinding(DefaultParameterSetName = 'NoQueueName')]
    param(

        $Connection,

        [parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$Exchange,

        [parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [ValidateSet('Direct','Fanout','Topic','Headers')]
        [string]$ExchangeType = $null,

        [parameter(ParameterSetName = 'NoQueueNameWithBasicQoS',Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [parameter(ParameterSetName = 'NoQueueName',Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [parameter(ParameterSetName = 'QueueName',Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
        [string[]]$Key,

        [parameter(ParameterSetName = 'QueueName',
                   Mandatory = $True,ValueFromPipelineByPropertyName = $true)]
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',
                   Mandatory = $True,ValueFromPipelineByPropertyName = $true)]
        [string]$QueueName,

        [parameter(ParameterSetName = 'QueueName',ValueFromPipelineByPropertyName = $true)]
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',ValueFromPipelineByPropertyName = $true)]
        [bool]$Durable = $true,

        [parameter(ParameterSetName = 'QueueName',ValueFromPipelineByPropertyName = $true)]
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',ValueFromPipelineByPropertyName = $true)]
        [bool]$Exclusive = $False,

        [parameter(ParameterSetName = 'QueueName',ValueFromPipelineByPropertyName = $true)]
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',ValueFromPipelineByPropertyName = $true)]
        [bool]$AutoDelete = $False,
        
        [parameter(parameterSetName = 'QueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [parameter(ParameterSetName = 'NoQueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [uint32]$prefetchSize,

        [parameter(parameterSetName = 'QueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [parameter(ParameterSetName = 'NoQueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [uint16]$prefetchCount,

        [parameter(parameterSetName = 'QueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [parameter(ParameterSetName = 'NoQueueNameWithBasicQoS',Mandatory = $true,ValueFromPipelineByPropertyName = $true)]
        [switch]$global
    )
    Try
    {
        $Channel = $Connection.CreateModel()

        Write-Progress -id 10 -Activity 'Create SCMB Connection' -Status 'Attempting connection to channel' -PercentComplete 80

        
            #if the ExchangeType is specified along with another property for non-default exchange ''
        if(  $ExchangeType -and ![string]::IsNullOrEmpty($Exchange) -and
               (![string]::IsNullOrEmpty($Durable) -or
               ![string]::IsNullOrEmpty($AutoDelete) )
        )
        {
            if([string]::IsNullOrEmpty($Durable)) {
                $Durable=$false
            }
            
            if([string]::IsNullOrEmpty($AutoDelete)){
                $AutoDelete = $false
            }

            #https://www.rabbitmq.com/releases/rabbitmq-dotnet-client/v3.6.6/rabbitmq-dotnet-client-3.6.6-client-htmldoc/html/
            #ExchangeDeclareNoWait(string exchange, string type, bool durable, bool autoDelete, IDictionary<string,object> arguments)
            #Actively declare the Exchange (as non-autodelete, non-durable)
            $ExchangeResult = $Channel.ExchangeDeclare($Exchange,$ExchangeType.ToLower(),$Durable,$AutoDelete,$null)
        }

        #Create a personal queue or bind to an existing queue
        if($QueueName)
        {
            $QueueResult = $Channel.QueueDeclare($QueueName, $Durable, $Exclusive, $AutoDelete, $null)
            if(-not $Key)
            {
                $Key = $QueueName
            }
        }
        else
        {
            if($PSEdition -eq 'core') { 
                #zero constructor is not impl in .Net Core lib
                $QueueResult = $Channel.QueueDeclare()
            }
        }
        $Arguments = [System.Collections.Generic.Dictionary[string,System.Object]]::new()
        if($PsCmdlet.ParameterSetName.Contains('BasicQoS')) {
            #Core version only has overload with 4 params.
            $channel.BasicQos($prefetchSize,$prefetchCount,$global,$Arguments)
        }
        #Bind our queue to the exchange
        foreach ($keyItem in $key) {
            if (![string]::IsNullOrEmpty($Exchange)) {
                $Channel.QueueBind($QueueName, $Exchange, $KeyItem, $Arguments)
            }
        }

        Write-Progress -id 10 -Activity 'Create SCMB Connection' -Status ('Conneccted to channel: {0}, {1}, {2}' -f $QueueName, $Exchange, $KeyItem) -PercentComplete 90

        $Channel
    }
    Catch
    {
        Throw $_
    }

}