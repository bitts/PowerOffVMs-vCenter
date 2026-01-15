@{
    # VMs que nunca devem ser desligadas
    ExcludedVMs = @(
        "vCenter01",
        "DC01",
        "DNS01"
    )

    # Delay entre desligamentos (segundos)
    DelayBetweenVMsSeconds = 15
}
