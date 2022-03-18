# Membership Expiration Changes  (MEC)

Tool to analyse the changes to the business_unit object to analyse changes in membership

From Microsoft Powershell use the following (requires Docker)
> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t mec -q https://github.com/iguzu/nbem.git) --for_month yyyy-mm <cbx_bu_change.xlsx> <results.xlsx>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t nbem -q https://github.com/iguzu/nbem.git) -h
> 