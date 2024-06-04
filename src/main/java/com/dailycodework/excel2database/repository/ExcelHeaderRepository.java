package com.dailycodework.excel2database.repository;

import com.dailycodework.excel2database.domain.ExcelHeader;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ExcelHeaderRepository extends JpaRepository<ExcelHeader, Long> {
}
