@{
    # The scripts accept -ClientSecret as [string] on purpose: the documented
    # usage passes it from an environment variable ($env:GRAPH_CLIENT_SECRET)
    # for CI/automation, and the docs steer production use toward certificate
    # auth (see docs/APP-REGISTRATION.md). A [SecureString] parameter would
    # break the env-var workflow without adding real protection in that model.
    # Everything else runs at Error + Warning severity.
    Severity     = @('Error','Warning')
    ExcludeRules = @(
        'PSAvoidUsingPlainTextForPassword',
        'PSAvoidUsingConvertToSecureStringWithPlainText'
    )
}
