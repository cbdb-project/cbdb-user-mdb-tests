# Precondition graph (.Enabled flips per event handler)

## LookAtEntry

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkUseYears_Click` | — | — | FrameYears, TxtFromYear, TxtToYear |
| `CmdAllPlaces_Click` | — | ChkSubUnits, ChkXYRef | — |
| `CmdImportEntryCodes_Click` | — | — | CmdQuery, CmdSaveEntryCodes |
| `CmdImportPlaces_Click` | ChkSubUnits, ChkXYRef | — | — |
| `CmdPickEntry_Click` | CmdQuery, CmdSaveEntryCodes | — | — |
| `CmdQuery_Click` | — | — | CmdGIS, CmdNeo4j, CmdStoreID |
| `CmdSelectPlace_Click` | ChkSubUnits, ChkXYRef, CmdAllPlaces | — | — |
| `FrameYears_Click` | — | TxtFromDynasty, TxtFromDynastyPY, TxtToDynasty, TxtToDynastyPY | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromYear, TxtToYear |

## LookAtKinship

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `CmdImport_Click` | — | — | CmdRun |
| `CmdRecallID_Click` | — | — | CmdRun |
| `CmdRun_Click` | — | — | ChkIncludeID, CmdGIS, CmdGUESS, CmdNeo4j, CmdStoreID, CmdUCINet, CmdUTF8Pajek |
| `CmdSelectPerson_Click` | — | — | CmdRun |
| `Form_Open` | CmdRecallID | — | — |

## LookAtOffice

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `CmdAllOffices_Click` | — | CmdAllOffices, CmdQuery | — |
| `CmdAllPlacesOffices_Click` | — | ChkSubUnitsOffice, ChkUseXY, CmdAllPlacesOffices, CmdQuery | — |
| `CmdAllPlacesPeople_Click` | — | ChkUseXY, CmdAllPlacesPeople, CmdQuery | — |
| `CmdImportOffices_Click` | — | — | CmdAllOffices, CmdQuery, CmdSaveOffices |
| `CmdImportPlaceOffice_Click` | ChkSubUnitsOffice, ChkUseXY | — | — |
| `CmdImportPlacePeople_Click` | ChkSubUnitsPeople, ChkUseXY | — | — |
| `CmdPickOffice_Click` | — | — | CmdAllOffices, CmdQuery, CmdSaveOffices |
| `CmdPlaceOffice_Click` | ChkSubUnitsOffice, ChkUseXY, CmdAllPlacesOffices, CmdQuery | — | — |
| `CmdPlacePeople_Click` | ChkUseXY, CmdAllPlacesPeople, CmdQuery | — | — |
| `CmdQuery_Click` | — | — | CmdGIS, CmdGISPeople, CmdNeo4j, CmdStoreID |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtOfficeFrom, TxtOfficeTo, TxtToDynasty, TxtToDynastyPY, TxtToYear |

## LookAtPlace

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkIndividual_Click` | — | — | CmdPickBAC |
| `CmdImportPlaces_Click` | ChkSubUnits, ChkXYRef, CmdQuery | — | — |
| `CmdQuery_Click` | — | — | CmdGephi, CmdNeo4j, CmdPajek, CmdStoreID, CmdUCINet |
| `CmdSelectPlace_Click` | ChkSubUnits, ChkXYRef | — | — |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtToDynasty, TxtToDynastyPY, TxtToYear |
| `QueryOK` | — | — | CmdQuery |

## LookAtAssociations

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkIndexYears_Click` | — | — | TxtFromYear, TxtToYear |
| `CmdAllPlaces_Click` | — | ChkSubUnits, ChkXYRef, FrameXY | — |
| `CmdImportAssociations_Click` | — | — | CmdQuery, CmdSaveAssociations |
| `CmdImportPlaces_Click` | ChkSubUnits, ChkXYRef, FrameXY | — | — |
| `CmdPickAssoc_Click` | — | — | CmdQuery, CmdSaveAssociations |
| `CmdQuery_Click` | — | — | CmdGIS, CmdGephi, CmdNeo4j, CmdPajek, CmdStoreID, CmdUCINet |
| `CmdSelectPlace_Click` | ChkSubUnits, ChkXYRef, CmdAllPlaces, FrameXY | — | — |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtToDynasty, TxtToDynastyPY, TxtToYear |

## LookAtAssociationPairs

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `CmdClearList_Click` | CmdImportList, CmdPickPerson1, CmdPickPerson2 | CmdClearList, CmdQuery | — |
| `CmdImportList_Click` | CmdClearList | CmdImportList, CmdPickPerson1, CmdPickPerson2 | CmdQuery |
| `CmdPickPerson1_Click` | CmdQuery | — | — |
| `CmdPickPerson2_Click` | CmdQuery | — | — |
| `CmdQuery_Click` | — | — | ChkIncludeID, CmdGIS, CmdGephi, CmdNeo4j, CmdPajek, CmdStoreID, CmdUCINet |
| `CmdRecallID_Click` | CmdClearList | CmdImportList, CmdPickPerson1, CmdPickPerson2 | CmdQuery |
| `Form_Open` | — | TxtFromYear, TxtToYear | CmdRecallID |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtToDynasty, TxtToDynastyPY, TxtToYear |

## LookAtStatus

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkIndexYears_Click` | — | — | TxtFromYear, TxtToYear |
| `CmdAllPlaces_Click` | — | ChkSubUnits, ChkXYRef, FrameXY | — |
| `CmdImportPlaces_Click` | ChkSubUnits, ChkXYRef, FrameXY | — | — |
| `CmdImportStatusCodes_Click` | — | — | CmdQuery, CmdSaveStatusCodes |
| `CmdPickStatus_Click` | — | — | CmdQuery, CmdSaveStatusCodes |
| `CmdQuery_Click` | — | — | CmdGIS, CmdNeo4j, CmdStoreID |
| `CmdSelectPlace_Click` | ChkSubUnits, ChkXYRef, CmdAllPlaces, FrameXY | — | — |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtToDynasty, TxtToDynastyPY, TxtToYear |

## LookAtTexts

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkIndexYears_Click` | — | — | TxtFromYear, TxtToYear |
| `CmdAllPlaces_Click` | — | ChkSubUnits, ChkXYRef, FrameXY | — |
| `CmdImportPlaces_Click` | ChkSubUnits, ChkXYRef, FrameXY | — | — |
| `CmdImportTextCategories_Click` | — | — | CmdQuery, CmdSaveTextCategories |
| `CmdPickTextCat_Click` | — | — | CmdQuery, CmdSaveTextCategories |
| `CmdQuery_Click` | — | — | CmdGIS, CmdNeo4j, CmdStoreID |
| `CmdSelectPlace_Click` | ChkSubUnits, ChkXYRef, CmdAllPlaces, FrameXY | — | — |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFromDynasty, TxtFromDynastyPY, TxtFromYear, TxtToDynasty, TxtToDynastyPY, TxtToYear |

## LookAtNetworks

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `CheckRunCriteria` | — | — | CmdRun |
| `ChkKin_Click` | — | TxtMaxCol, TxtMaxDwn, TxtMaxMar, TxtMaxUp | ChkKinshipParam |
| `ChkKinshipParam_Click` | — | — | TxtMaxCol, TxtMaxDwn, TxtMaxMar, TxtMaxUp |
| `ChkNonKin_Click` | — | — | ChkFamily, ChkFinance, ChkFriendship, ChkMedicine, ChkMilitaryAll, ChkMilitaryOppose, ChkMilitarySupport, ChkPolEqual, ChkPolOppose, ChkPolSponsor, ChkPolSub, ChkPolSup, ChkPolSupport, ChkPoliticsAll, ChkReligion, ChkSchAffiliation, ChkSchAttack, ChkSchLitArt, ChkSchMember, ChkSchPatron, ChkSchTeacher, ChkSchTopic, ChkScholarshipAll, ChkSelectAll, ChkWriBiog, ChkWriCommem, ChkWriEpitaph, ChkWriExplain, ChkWriLetters, ChkWriMottos, ChkWriOccasion, ChkWriPreface, ChkWriRitual, ChkWritingsAll |
| `CmdAllPeople_Click` | — | CmdAllPeople, CmdRerun | — |
| `CmdAllPlaces_Click` | — | ChkPlaceLimit, ChkSubUnits, ChkXYRef | — |
| `CmdImportPeople_Click` | CmdAllPeople, CmdRun | CmdRerun | — |
| `CmdImportPlaces_Click` | ChkPlaceLimit, ChkSubUnits, ChkXYRef, CmdRun | — | — |
| `CmdRecallID_Click` | CmdRerun | — | CmdAllPeople, CmdRun |
| `CmdRun_Click` | — | — | ChkIncludeID, CmdGIS, CmdGUESS, CmdNeo4j, CmdPajek, CmdRerun, CmdStoreID, CmdUCINet |
| `CmdSelectPerson_Click` | CmdAllPeople, CmdRun | CmdRerun | — |
| `CmdSelectPlace_Click` | ChkPlaceLimit, ChkSubUnits, ChkXYRef, CmdAllPlaces, CmdRun | — | — |
| `Form_Open` | CmdRecallID | ChkPlaceLimit, CmdAllPlaces | — |
| `FrameFilterYears_Click` | — | — | CmdAllDynasties, CmdFromDynasty, CmdToDynasty, TxtFrom, TxtFromDynasty, TxtFromDynastyPY, TxtTo, TxtToDynasty, TxtToDynastyPY |

## LookAtGroupData

| Event handler | Enables | Disables | Toggles |
|---|---|---|---|
| `ChkAddr_Click` | — | — | CmdRun, FrameQueryAddress |
| `ChkEntry_Click` | — | — | CmdRun |
| `ChkOffice_Click` | — | — | CmdRun |
| `ChkStatus_Click` | — | — | CmdRun |
| `ChkText_Click` | — | — | CmdRun |
| `CmdImport_Click` | — | — | CmdRun, CmdStoreID |
| `CmdRecallID_Click` | — | — | CmdRun, CmdStoreID |
| `CmdRun_Click` | — | — | CmdNeo4j |
| `Form_Open` | CmdRecallID | — | — |
| `queryAddr` | — | — | ChkGisAddr, CmdGIS |
| `queryEntry` | — | — | ChkGisEntry, CmdGIS |
| `queryOffice` | — | — | ChkGisOffice, ChkGisOfficePeople, CmdGIS |
| `queryStatus` | — | — | ChkGisStatus, CmdGIS |
| `queryText` | — | — | ChkGisText, CmdGIS |
