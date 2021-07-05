Configuration DependsOnExample {
    Node Test-PC1 {
        Group GroupExample {
            Ensure = "Present"
            GroupName = "TestGroup"
        }

        User UserExample {
            Ensure = "Present"
            UserName = "TestUser"
            FullName = "TestUser"
            DependsOn = "[Group]GroupExample"
        }
    }
}

DependsOnExample
