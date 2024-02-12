# ...

# Register events
$event = Register-ObjectEvent $watcher "Created" -Action $action

function Stop-Watcher {
    Unregister-Event -SourceIdentifier $event.Name
    $watcher.Dispose()
    Write-Host "Watcher stopped."
}

        Stop-Watcher
 
