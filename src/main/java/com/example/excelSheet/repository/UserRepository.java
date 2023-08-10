package com.example.excelSheet.repository;

import com.example.excelSheet.model.UserActionLog;
import com.example.excelSheet.model.UserActionLogDTO;
import jakarta.transaction.Transactional;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.math.BigInteger;
import java.util.List;
@Repository
public interface UserRepository extends JpaRepository<UserActionLog, String> {
    @Query(value = "SELECT * FROM user_action_log", nativeQuery = true)
    List<UserActionLog> findAllUsingNativeQuery();
    @Query(value = "SELECT * FROM user_action_log WHERE action=?1", nativeQuery = true)
    List<UserActionLog> findAllUsingNativeQuery(String action);
    @Modifying
    @Transactional
    @Query(value = "INSERT INTO user_action_log(id, email, type, action, created_timestamp, message,status) values (?1,?2,?3,?4,?5,?6,?7)", nativeQuery = true)
    void addUserAction(String id, String email, String type, String action, BigInteger createdTimestamp, String message, String status);

    @Query(value = "SELECT *"
            + " FROM user_action_log ual"
            + " WHERE (?1 = 'all' or ual.email = ?1) AND (?2 = 'all' or ual.type = ?2) AND (?3 = 'all' or ual.action = ?3) "
            + " AND (?4 IS NULL or ual.created_timestamp >= ?4) AND (?5 IS NULL or ual.created_timestamp <= ?5) AND (?6 = 'all' or ual.status = ?6) "
            + " AND (?7 = 'null'  or CONCAT_WS('', ual.email, ual.type, ual.action, ual.message, ual.status) LIKE CONCAT('%',?7,'%')) "
            + " AND (?8 = 'all' or CONCAT_WS('', ual.message) LIKE CONCAT('%',?8,'%'))"
            + " ORDER BY ual.created_timestamp DESC, ual.id ",nativeQuery = true)
    List<UserActionLogDTO> getUserActionLogs(String email, String type, String action, BigInteger start_date, BigInteger end_date, String status, Integer pagesize, Integer offset, String searchkey, String message);

    //Get user action log records count
    @Query(value = "SELECT COUNT(*) FROM user_action_log",nativeQuery = true)
    Integer getTotalUserActionLogRecords();

    //Delete  user action log records
    @Modifying
    @Transactional
    @Query(value = "DELETE FROM user_action_log ORDER BY created_timestamp ASC LIMIT ?1",nativeQuery = true)
    void deleteUserActionLogRecords(Integer limit);

    @Query(value = "SELECT COUNT(*)"
            + " FROM user_action_log ual"
            + " WHERE (?1 = 'all' or ual.email = ?1) AND (?2 = 'all' or ual.type = ?2) AND (?3 = 'all' or ual.action = ?3) "
            + " AND (?4 IS NULL or ual.created_timestamp >= ?4) AND (?5 IS NULL or ual.created_timestamp <= ?5) AND (?6 = 'all' or ual.status = ?6) "
            + " AND (?7 = 'null'  or CONCAT_WS('', ual.email, ual.type, ual.action, ual.message, ual.status) LIKE CONCAT('%',?7,'%')) "
            + " AND (?8 = 'all' or CONCAT_WS('', ual.message) LIKE CONCAT('%',?8,'%'))"
            + " ORDER BY ual.created_timestamp DESC, ual.id ",nativeQuery = true)
    Integer getUserActionLogsCount(String email, String type, String action, BigInteger start_date, BigInteger end_date, String status, String searchkey, String message);
}

