-- Separate reviewed phase, ACCOUNTADMIN only. Parent first checks schema exact owner,
-- managed-access/comment, empty contents and NO procedure of this name/any overload.
-- Stop on error. Scalar-only: no tables, writeback, owner execution, or dynamic SQL.
CREATE PROCEDURE SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO(
    P_OPERATION_ID VARCHAR,
    P_AMOUNT NUMBER(18,2),
    P_COUNT NUMBER(38,0),
    P_ENABLED BOOLEAN,
    P_OCCURRED_AT TIMESTAMP_NTZ(3)
)
RETURNS VARCHAR
LANGUAGE SQL
CALLED ON NULL INPUT
COMMENT = 'Task5 reviewed scalar-only procedure fixture; contract v1'
EXECUTE AS CALLER
AS
$$
DECLARE
  required_input EXCEPTION (-20051, 'TASK5_REQUIRED_INPUT');
  invalid_operation EXCEPTION (-20052, 'TASK5_INVALID_OPERATION_ID');
  negative_count EXCEPTION (-20053, 'TASK5_NEGATIVE_COUNT');
BEGIN
  IF (P_OPERATION_ID IS NULL OR P_AMOUNT IS NULL OR P_COUNT IS NULL
      OR P_ENABLED IS NULL OR P_OCCURRED_AT IS NULL) THEN
    RAISE required_input;
  END IF;
  IF (LENGTH(TRIM(P_OPERATION_ID)) = 0 OR CONTAINS(P_OPERATION_ID, '|')) THEN
    RAISE invalid_operation;
  END IF;
  IF (P_COUNT < 0) THEN
    RAISE negative_count;
  END IF;
  RETURN 'operation_id=' || P_OPERATION_ID
      || '|amount=' || TRIM(TO_VARCHAR(P_AMOUNT, '9999999999999990.00'))
      || '|count=' || TRIM(TO_VARCHAR(P_COUNT, '99999999999999999999999999999999999990'))
      || '|enabled=' || IFF(P_ENABLED, 'true', 'false')
      || '|occurred_at=' || TO_VARCHAR(P_OCCURRED_AT, 'YYYY-MM-DD"T"HH24:MI:SS.FF3');
END;
$$;
-- Parent captures DESCRIBE, GET_DDL and SHOW GRANTS; reconcile normalized Snowflake
-- type identity versus declaration before granting. Signature is NOT Sigma display text.
DESCRIBE PROCEDURE SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO(VARCHAR, NUMBER, NUMBER, BOOLEAN, TIMESTAMP_NTZ);
SELECT GET_DDL('PROCEDURE', 'SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO(VARCHAR, NUMBER, NUMBER, BOOLEAN, TIMESTAMP_NTZ)');
SHOW PROCEDURES LIKE 'TASK5_TYPED_ECHO' IN SCHEMA SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST;
