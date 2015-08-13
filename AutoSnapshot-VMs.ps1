# Create and rotate Hyper-V Snapshots
# Pulled from - https://support.managed.com/kb/a1872/how-to-automatically-rotate-hyperv-snapshots.aspx

$Days = 7

$VMs = Get-VM
foreach($VM in $VMs){
    $Snapshots = Get-VMSnapshot $VM
    foreach($Snapshot in $Snapshots){
    
        if ($snapshot.CreationTime.AddDays($Days) -lt (get-date)){
            Remove-VMSnapshot $Snapshot
        } 
    }
    
    Checkpoint-VM $VM
}
