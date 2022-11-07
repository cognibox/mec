# Membership Expiration Changes  (MEC)

Tool to analyse the changes to the business_unit object to analyse changes in membership
Query In redash to use
SBL - BU Object Changes

From Microsoft Powershell use the following (requires Docker)
> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t mec -q https://github.com/cognibox/mec.git) --for_month yyyy-mm <cbx_bu_change.xlsx> <results.xlsx>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t nbem -q https://github.com/cognibox/mec.git) -h

The tool will classify membership changes in the following way:

1. Expiration that moves forward more than 24 months are considered  "free accounts"
2. Expirations that moves less than 3 months in the future are consider "change_expiration" (i.e. not really a change)
3. Expirations that have no prior expiration date or that the prior expiration date is 6 months prior, are considered "new accounts"
4. All others are "renewal accounts"
