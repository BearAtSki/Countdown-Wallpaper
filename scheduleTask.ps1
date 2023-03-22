param ($countDownBG)

$action = New-ScheduledTaskAction -Execute $countDownBG

$trigger1 =  New-ScheduledTaskTrigger -AtLogon
$trigger2 =  New-ScheduledTaskTrigger -Daily -At 12:00:01am

Register-ScheduledTask -Action $action -Trigger $trigger1,$trigger2 -TaskName "CountdownWP" -Description "Daily countdown wallpaper"