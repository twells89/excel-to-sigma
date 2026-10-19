-- Reviewed v4 reference source; parent owns provisioning and actual evidence.
-- Offline fixture only: never automatically execute or replace an existing object.
-- Execute only after fresh name/overload absence and unchanged original/schema grants.
CREATE PROCEDURE SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO_TEXT_AMOUNT(
    P_OPERATION_ID VARCHAR,
    P_AMOUNT_TEXT VARCHAR,
    P_COUNT NUMBER(38,0),
    P_ENABLED BOOLEAN,
    P_OCCURRED_AT TIMESTAMP_NTZ(3)
)
RETURNS VARCHAR
LANGUAGE SQL
CALLED ON NULL INPUT
COMMENT = 'Task5 explicit decimal-text transport adapter; delegates unchanged typed business contract; v4'
EXECUTE AS CALLER
AS
$$
DECLARE
  required_input EXCEPTION (-20051, 'TASK5_REQUIRED_INPUT');
  invalid_amount_text EXCEPTION (-20054, 'TASK5_INVALID_AMOUNT_TEXT');
  typed_amount NUMBER(18,2);
  scalar_result VARCHAR;
BEGIN
  IF (P_OPERATION_ID IS NULL OR P_AMOUNT_TEXT IS NULL OR P_COUNT IS NULL
      OR P_ENABLED IS NULL OR P_OCCURRED_AT IS NULL) THEN
    RAISE required_input;
  END IF;
  IF (REGEXP_INSTR(P_AMOUNT_TEXT, '[[:space:]]') > 0
      OR NOT REGEXP_LIKE(P_AMOUNT_TEXT, '^-?(0|[1-9][0-9]{0,14})([.][0-9]{1,2})?$')) THEN
    RAISE invalid_amount_text;
  END IF;
  typed_amount := TRY_TO_DECIMAL(P_AMOUNT_TEXT, 18, 2);
  IF (typed_amount IS NULL OR ABS(typed_amount) >= 1000000000000000
      OR (typed_amount = 0 AND LEFT(P_AMOUNT_TEXT, 1) = '-')) THEN
    RAISE invalid_amount_text;
  END IF;
  CALL SIGMA_WRITE_DB.SIGMA_TASK5_SP_TEST.TASK5_TYPED_ECHO(
      :P_OPERATION_ID, :typed_amount, :P_COUNT, :P_ENABLED, :P_OCCURRED_AT
  ) INTO :scalar_result;
  RETURN scalar_result;
END;
$$;
