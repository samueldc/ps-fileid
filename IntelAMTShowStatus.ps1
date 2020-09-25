Set-ExecutionPolicy remotesigned

Import-Module IntelvPro

$8021XProfileRef =$wsmanConnectionObject.NewReference("SELECT * FROM AMT_8021XProfile WHERE InstanceID='Intel(r) AMT 802.1x Profile 0'")

$8021XProfileInstance =$8021XProfileRef.Get()

$enabled =$8021XProfileInstance.GetProperty("Enabled")

$authenticationProtocol =$8021XProfileInstance.GetProperty("AuthenticationProtocol")

$domain =$8021XProfileInstance.GetProperty("Domain")