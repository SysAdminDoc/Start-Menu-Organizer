@{
    Severity = @('Error', 'Warning')

    # This is a single-file WPF app, not an exported cmdlet module. Local
    # event-handler functions intentionally use UI/action names instead of
    # public PowerShell verb-noun API names and ShouldProcess contracts.
    ExcludeRules = @(
        'PSUseApprovedVerbs',
        'PSUseSingularNouns',
        'PSUseShouldProcessForStateChangingFunctions',
        'PSReviewUnusedParameter'
    )
}
