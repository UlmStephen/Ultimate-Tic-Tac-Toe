-- --------------------------------------------------------------------------------
-- Name: dbUltimateTicTacToe
-- Author: Stephen Ulm
-- Class: Capstone
-- --------------------------------------------------------------------------------
USE dbUltimateTicTacToe
-- --------------------------------------------------------------------------------
-- Drop tables
-- --------------------------------------------------------------------------------

IF OBJECT_ID( 'TGames' )	IS NOT NULL DROP TABLE TGames
IF OBJECT_ID( 'TGameDifficulties' )	IS NOT NULL DROP TABLE TGameDifficulties
IF OBJECT_ID( 'TGameOutcomes' )	IS NOT NULL DROP TABLE TGameOutcomes
IF OBJECT_ID( 'TGameMoves' )	IS NOT NULL DROP TABLE TGameMoves
SELECT * FROM TGameMoves
-- --------------------------------------------------------------------------------
-- Step  .1: Normalize and Create Tables
-- --------------------------------------------------------------------------------

CREATE TABLE TGames
(
	 intGameID						INTEGER			NOT NULL		IDENTITY(1,1)
	,strPlayerName					VARCHAR(50)		NOT NULL
	,blnComputerMovesFirst			BIT				NOT NULL
	,intGameDifficultyID			INTEGER			NOT NULL
	,intGameOutcomeID				INTEGER			NOT NULL
	,dtmPlayed						DATETIME		NOT NULL
	,CONSTRAINT TGames_PK PRIMARY KEY ( intGameID ) 
)
CREATE TABLE TGameDifficulties
(
	 intGameDifficultyID			INTEGER			NOT NULL		IDENTITY(1,1)
	,strGameDifficulty				VARCHAR(50)		NOT NULL
	,CONSTRAINT TGameDifficulties_PK PRIMARY KEY ( intGameDifficultyID ) 
)
CREATE TABLE TGameOutcomes
(
	 intGameOutcomeID				INTEGER			NOT NULL		IDENTITY(1,1)
	,strGameOutcome					VARCHAR(50)		NOT NULL
	,CONSTRAINT TGameOutcomes_PK PRIMARY KEY ( intGameOutcomeID ) 
)
CREATE TABLE TGameMoves
(
	 intGameID						INTEGER			NOT NULL		
	,intMoveIndex					INTEGER			NOT NULL
	,intGroup						INTEGER			NOT NULL
	,intSquare						INTEGER			NOT NULL
	,CONSTRAINT TGameMoves_PK PRIMARY KEY ( intGameID, intMoveIndex ) 
)


-- --------------------------------------------------------------------------------
-- Step .2: Make Foreign Keys
-- --------------------------------------------------------------------------------
--# Child				Parent				Column(s)
--1	TGames				TGameDifficulties	intGameDifficultyID
--2	TGames				TGameOutcomes		intGameOutcomeID
--3	TGameMoves			TGames				intGameID
-- --------------------------------------------------------------------------------

--1
ALTER TABLE TGames ADD CONSTRAINT TGames_TGameDifficulties_FK
FOREIGN KEY ( intGameDifficultyID ) REFERENCES TGameDifficulties( intGameDifficultyID )

--2
ALTER TABLE TGames ADD CONSTRAINT TGames_TGameOutcomes_FK
FOREIGN KEY ( intGameOutcomeID ) REFERENCES TGameOutcomes( intGameOutcomeID )

--3
ALTER TABLE TGameMoves ADD CONSTRAINT TGameMoves_TGames_FK
FOREIGN KEY ( intGameID ) REFERENCES TGames( intGameID )


-- --------------------------------------------------------------------------------
-- Step .3: Add Sample Data
-- --------------------------------------------------------------------------------

INSERT INTO TGameDifficulties ( strGameDifficulty ) 
VALUES ('Easy')

INSERT INTO TGameDifficulties ( strGameDifficulty ) 
VALUES ('Hard')
SELECT * FROM TGameOutcomes
INSERT INTO TGameOutcomes ( strGameOutcome ) 
VALUES ( 'Player Won')

INSERT INTO TGameOutcomes ( strGameOutcome ) 
VALUES ( 'Computer Won')

INSERT INTO TGameOutcomes ( strGameOutcome ) 
VALUES ( 'Draw')

INSERT INTO TGameOutcomes ( strGameOutcome ) 
VALUES ( 'Game Ended Early')

GO
DELETE FROM TGames
GO
CREATE PROCEDURE uspAddMove
	 @intGameID AS INTEGER
	,@intMoveIndex AS INTEGER
	,@intGroup AS INTEGER
	,@intSquare AS INTEGER
AS
SET NOCOUNT ON	-- Report only errors
SET XACT_ABORT ON	-- Terminate and rollback entire transaction on error
 
BEGIN TRANSACTION

	INSERT INTO TGameMoves
	VALUES(@intGameID, @intMoveIndex, @intGroup, @intSquare)

COMMIT TRANSACTION
GO

SELECT 
	TG.intGameID
	,TGO.strGameOutcome
FROM 
	 TGames AS TG
	 ,TGameOutcomes AS TGO
WHERE
	TG.intGameOutcomeID = TGO.intGameOutcomeID

	SELECT * FROM TGameMoves


GO

CREATE VIEW VCompletedGames
AS
SELECT 
	 TG.intGameID
	,TG.strPlayerName
	,TGO.strGameOutcome
	,Count(TGM.intMoveIndex) AS intMoveCount

FROM 
	  TGames AS TG
	 ,TGameOutcomes AS TGO
	 ,TGameMoves AS TGM
WHERE
	TG.intGameOutcomeID = TGO.intGameOutcomeID
AND TG.intGameID = TGM.intGameID
AND TGO.intGameOutcomeID != 4 --only include completed ideas
GROUP BY
	 TG.intGameID
	,TG.strPlayerName
	,TGO.strGameOutcome


GO

CREATE VIEW VGameReplayInformation
AS
SELECT 
	 TG.intGameID
	,TG.strPlayerName
	,TG.blnComputerMovesFirst
	,TGM.intMoveIndex
	,TGM.intGroup	
	,TGM.intSquare	
FROM 
	  TGames AS TG
	 ,TGameMoves AS TGM
WHERE
	TG.intGameID = TGM.intGameID

GROUP BY
	 TG.intGameID
	,TG.strPlayerName
	,TG.blnComputerMovesFirst
	,TGM.intMoveIndex
	,TGM.intGroup	
	,TGM.intSquare	


GO
