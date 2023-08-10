package com.example.excelSheet.model;


import jakarta.persistence.*;
import java.math.BigInteger;

@SqlResultSetMapping(
        name = "userActionLogMapping",
        classes = {
                @ConstructorResult(
                        targetClass = UserActionLogDTO.class,
                        columns = {
                                @ColumnResult(name = "id", type = String.class),
                                @ColumnResult(name = "email", type = String.class),
                                @ColumnResult(name = "type", type = String.class),
                                @ColumnResult(name = "action", type = String.class),
                                @ColumnResult(name = "created_timestamp", type = BigInteger.class),
                                @ColumnResult(name = "message", type = String.class),
                                @ColumnResult(name = "status", type = String.class)

                        }
                )
        }
)

// getUserActionLogs
@NamedNativeQuery(
        name = "UserActionLog.getUserActionLogs",
        query = "SELECT  ual.id, ual.email, ual.type, ual.action, ual.created_timestamp, ual.message, ual.status "
                + " FROM user_action_log ual"
                + " WHERE (?1 = 'all' or ual.email = ?1) AND (?2 = 'all' or ual.type = ?2) AND (?3 = 'all' or ual.action = ?3) "
                + " AND (?4 IS NULL or ual.created_timestamp >= ?4) AND (?5 IS NULL or ual.created_timestamp <= ?5) AND (?6 = 'all' or ual.status = ?6) "
                + " AND (?9 = 'null'  or CONCAT_WS('', ual.email, ual.type, ual.action, ual.message, ual.status) LIKE CONCAT('%',?9,'%'))"
                + " AND (?10 = 'all' or CONCAT_WS('', ual.message) LIKE CONCAT('%',?10,'%'))"
                + " ORDER BY ual.created_timestamp DESC, ual.id "
                + " LIMIT ?7 OFFSET ?8",
        resultSetMapping = "userActionLogMapping"
)
@Entity
public class UserActionLog {

    @Id
    private String id;

    @Column(length = 128)
    private String email;

    @Column(length = 128)
    private String type;

    @Column(length = 128)
    private String action;

    @Column(length = 128)
    private BigInteger created_timestamp;

    @Column(columnDefinition = "TEXT")
    private String message;

    @Column(length = 128)
    private String status;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getAction() {
        return action;
    }

    public void setAction(String action) {
        this.action = action;
    }

    public BigInteger getCreated_timestamp() {
        return created_timestamp;
    }

    public void setCreated_timestamp(BigInteger created_timestamp) {
        this.created_timestamp = created_timestamp;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }
}

