package com.example.excelSheet.controller;

import com.alibaba.fastjson.JSONObject;
import com.example.excelSheet.model.UserActionLogDTO;
import com.example.excelSheet.service.UserActionLogService;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/products")
public class Controller {

    private final UserActionLogService productService;
    @Autowired
    public Controller(UserActionLogService productService) {
        this.productService = productService;
    }

    @Autowired
    UserActionLogService userActionLogService;

    // Get User Action Logs
    @RequestMapping(method = RequestMethod.POST, value = "/user/{username}/vdms/{vdmsid}/getuseractionlogs")
    public List<UserActionLogDTO> getUserActionLogs(@PathVariable String username, @PathVariable String vdmsid,
                                                    @RequestParam(defaultValue = "1") Integer pageno,
                                                    @RequestParam(defaultValue = "10") Integer pagesize,
                                                    @RequestParam(defaultValue = "null") String searchkey,
                                                    @RequestBody JSONObject filterObject){
        return userActionLogService.getUserActionLogs(username, vdmsid, filterObject, pageno, pagesize, searchkey);
    }

    // Get User Action Logs Count
    @RequestMapping(method = RequestMethod.POST, value = "/user/{username}/vdms/{vdmsid}/getuseractionlogscount")
    public Map<String,Integer> getUserActionLogsCount(@PathVariable String username, @PathVariable String vdmsid,
                                                      @RequestParam(defaultValue = "null") String searchkey,
                                                      @RequestBody JSONObject filterObject){
        return userActionLogService.getUserActionLogsCount(username, vdmsid, filterObject,searchkey);
    }

    @GetMapping("/export-to-excel")
    public ResponseEntity<byte[]> exportToExcel(@PathVariable String username, @PathVariable String vdmsid,
                                                @RequestParam(defaultValue = "1") Integer pageno,
                                                @RequestParam(defaultValue = "10") Integer pagesize,
                                                @RequestParam(defaultValue = "null") String searchkey,
                                                @RequestBody JSONObject filterObject,
                                                HttpServletResponse response) throws Exception {
        return productService.getAllProductsUsingNativeQuery(username, vdmsid, filterObject, pageno, pagesize, searchkey);

    }
}
