UPDATE
  (
    OLCCAgents
    INNER JOIN Agents ON OLCCAgents.AgentNumber = Agents.AgentNumber
  )
  INNER JOIN OLCCAgentsXLS ON OLCCAgents.AgentNumber = OLCCAgentsXLS.F1
SET
  Agents.FirstName = [OLCCAgentsXLS]![F6],
  Agents.LastName = [OLCCAgentsXLS]![F2],
  Agents.Address = Left(
    [OLCCAgentsXLS].[F9],
    (
      Len([OLCCAgentsXLS].[f9])-5
    )
  ),
  Agents.Zip = Right([OLCCAgentsXLS].[f9], 5),
  Agents.Phone = [OLCCAgentsXLS]![F3],
  Agents.ChangeDate = Now();
