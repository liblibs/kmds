INSERT INTO Agents (
  AgentNumber, LastName, FirstName,
  Phone, Address, City, Zip, ST
)
SELECT
  OLCCAgentsXLS.F1,
  OLCCAgentsXLS.F2,
  OLCCAgentsXLS.F6,
  OLCCAgentsXLS.F3,
  Left(
    [F9],
    (
      Len([f9])-5
    )
  ) AS addy,
  OLCCAgentsXLS.F2,
  Right([f9], 5) AS zp,
  "OR" AS state
FROM
  Agents
  RIGHT JOIN OLCCAgentsXLS ON Agents.AgentNumber = OLCCAgentsXLS.F1
WHERE
  (
    (
      (Agents.AgentNumber) Is Null
    )
  );
