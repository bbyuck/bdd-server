package com.bb.bdd.domain.excel.repository;

import com.bb.bdd.domain.excel.ShopCode;
import com.bb.bdd.domain.excel.entity.Item;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface ItemRepository extends JpaRepository<Item, Long> {

    List<Item> findItemsByShop(ShopCode shop);
}
