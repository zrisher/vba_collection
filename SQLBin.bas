Attribute VB_Name = "SQLBin"
'-- -----------------------------------------------------
'-- Table main.assessments
'-- -----------------------------------------------------
'CREATE TABLE main.assessments (
'  id INTEGER PRIMARY KEY,
'  topic_id INTEGER NOT NULL,
'  rating_id INTEGER,
'  explanation TEXT,
'  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  CONSTRAINT fk_topics_assessments FOREIGN KEY
'    (topic_id)
'    REFERENCES topics(id)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT,
'  CONSTRAINT fk_ratings_assessments FOREIGN KEY
'    (rating_id)
'    REFERENCES ratings(id)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT
');
'
'CREATE INDEX main.ix_assessments_topic_id
'   ON assessments(topic_id);
'CREATE INDEX main.ix_assessments_rating_id
'   ON assessments(rating_id);
'
'CREATE TRIGGER main.updatetimestamp_assessments
'   AFTER UPDATE
'   ON assessments
'   FOR EACH ROW
'   BEGIN
'       UPDATE assessments
'       SET updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.candidates
'-- -----------------------------------------------------
'CREATE TABLE candidates (
'   id INTEGER PRIMARY KEY,
'   name TEXT NOT NULL,
'   requisition_id INTEGER,
'   created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   CONSTRAINT fk_requisitions_candidates FOREIGN KEY
'       (requisition_id)
'       REFERENCES requisitions(id)
'       ON DELETE RESTRICT
'       ON UPDATE RESTRICT
');
'CREATE INDEX ix_candidates_requisition_id
'   ON candidates(requisition_id);
'
'CREATE TRIGGER updatetimestamp_candidates
'   AFTER Update
'   ON candidates
'   FOR EACH ROW
'   BEGIN
'       Update candidates
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.interviews
'-- -----------------------------------------------------
'CREATE TABLE interviews (
'   id INTEGER PRIMARY KEY,
'   interviewer_id INTEGER NOT NULL,
'   candidate_id INTEGER NOT NULL,
'   starttime DATETIME NOT NULL,
'   duration DATETIME NOT NULL,
'   stage_id INTEGER,
'   location_id INTEGER,
'   created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   CONSTRAINT fk_employees_interviews FOREIGN KEY
'       (interviewer_id)
'       REFERENCES employees(id)
'       ON DELETE RESTRICT
'       ON UPDATE RESTRICT,
'   CONSTRAINT fk_candidates_interviews FOREIGN KEY
'       (candidate_id)
'       REFERENCES candidates(id)
'       ON DELETE RESTRICT
'       ON UPDATE RESTRICT,
'   CONSTRAINT fk_stages_interviews FOREIGN KEY
'       (stage_id)
'       REFERENCES stages(id)
'       ON DELETE RESTRICT
'       ON UPDATE RESTRICT,
'   CONSTRAINT fk_locations_interviews FOREIGN KEY
'       (location_id)
'       REFERENCES locations(id)
'       ON DELETE RESTRICT
'       ON UPDATE RESTRICT,
');
'CREATE INDEX ix_interviews_interviewer_id
'   ON interviews(interviewer_id);
'CREATE INDEX ix_interviews_candidate_id
'   ON interviews(candidate_id);
'CREATE INDEX ix_interviews_stage_id
'   ON interviews(stage_id);
'CREATE INDEX ix_interviews_location_id
'   ON interviews(location_id);
'
'CREATE TRIGGER updatetimestamp_interviews
'   AFTER Update
'   ON interviews
'   FOR EACH ROW
'   BEGIN
'       Update interviews
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.locations
'-- -----------------------------------------------------
'CREATE TABLE locations (
'   id INTEGER PRIMARY KEY,
'   name TEXT NOT NULL,
'   created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
');
'CREATE TRIGGER updatetimestamp_locations
'   AFTER Update
'   ON locations
'   FOR EACH ROW
'   BEGIN
'       Update locations
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.ratings
'-- -----------------------------------------------------
'CREATE TABLE main.ratings (
'  id INTEGER PRIMARY KEY,
'  name TEXT NOT NULL,
'  ratingtype_id INTEGER NOT NULL,
'  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  CONSTRAINT fk_ratingtypes_ratings FOREIGN KEY
'    (ratingtype_id)
'    REFERENCES ratingtypes(ID)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT
');
'
'CREATE INDEX main.ix_ratings_ratingtype_id
'   ON ratings(ratingtype_id);
'
'CREATE TRIGGER main.updatetimestamp_ratings
'   AFTER Update
'   ON ratings
'   FOR EACH ROW
'   BEGIN
'       Update ratings
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.ratingtypes
'-- -----------------------------------------------------
'CREATE TABLE main.ratingtypes (
'  id INTEGER PRIMARY KEY,
'  name TEXT NOT NULL,
'  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
');
'
'CREATE TRIGGER main.updatetimestamp_ratingtypes
'   AFTER Update
'   ON ratingtypes
'   FOR EACH ROW
'   BEGIN
'       Update ratingtypes
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.requisitions
'-- -----------------------------------------------------
'CREATE TABLE requisitions (
'   id INTEGER PRIMARY KEY,
'   name TEXT NOT NULL,
'   created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
');
'
'CREATE TRIGGER updatetimestamp_requisitions
'   AFTER Update
'   ON requisitions
'   FOR EACH ROW
'   BEGIN
'       Update requisitions
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.stages
'-- -----------------------------------------------------
'CREATE TABLE stages (
'   id INTEGER PRIMARY KEY,
'   name TEXT NOT NULL,
'   created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'   updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
');
'CREATE TRIGGER updatetimestamp_stages
'   AFTER Update
'   ON stages
'   FOR EACH ROW
'   BEGIN
'       Update stages
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;

'-- -----------------------------------------------------
'-- Table main.topics
'-- -----------------------------------------------------
'CREATE TABLE main.topics (
'  id INTEGER PRIMARY KEY,
'  name TEXT NOT NULL,
'  topictype_id INTEGER NOT NULL,
'  ratingtype_id INTEGER NOT NULL,
'  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  CONSTRAINT fk_topictypes_topics FOREIGN KEY
'    (topictype_id)
'    REFERENCES topictypes(id)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT,
'  CONSTRAINT fk_ratingtypes_topics FOREIGN KEY
'    (ratingtype_id)
'    REFERENCES ratingtypes(id)
'    ON DELETE RESTRICT
'    ON UPDATE RESTRICT
');
'
'CREATE INDEX main.ix_topics_topictype_id
'   ON topics(topictype_id);
'CREATE INDEX main.ix_topics_ratingtype_id
'   ON topics(ratingtype_id);
'
'CREATE TRIGGER main.updatetimestamp_topics
'   AFTER UPDATE
'   ON topics
'   FOR EACH ROW
'   BEGIN
'       UPDATE topics
'       SET updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;


'-- -----------------------------------------------------
'-- Table main.topictypes
'-- -----------------------------------------------------
'CREATE TABLE main.topictypes (
'  id INTEGER PRIMARY KEY,
'  name TEXT NOT NULL,
'  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
'  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
');
'
'CREATE TRIGGER main.updatetimestamp_topictypes
'   AFTER Update
'   ON topictypes
'   FOR EACH ROW
'   BEGIN
'       Update topictypes
'       Set updated_at = CURRENT_TIMESTAMP
'       WHERE id = OLD.id;
'   End;
