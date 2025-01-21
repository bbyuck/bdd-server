package com.bb.bdd.domain.excel.repository;

import com.bb.bdd.domain.excel.entity.Item;
import org.springframework.data.jpa.repository.JpaRepository;

public interface ItemRepository extends JpaRepository<Item, Long> {
}
