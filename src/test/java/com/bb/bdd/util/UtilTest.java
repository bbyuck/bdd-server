package com.bb.bdd.util;

import com.bb.bdd.domain.excel.ShopCode;
import com.bb.bdd.domain.excel.entity.Item;
import com.bb.bdd.domain.excel.repository.ItemRepository;
import com.bb.bdd.domain.excel.service.ExcelTransformService;
import com.bb.bdd.domain.excel.util.ExcelReader;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;

@SpringBootTest
public class UtilTest {


    @Autowired
    ExcelReader excelReader;

    @Autowired
    ItemRepository itemRepository;

    @Autowired
    ExcelTransformService excelTransformService;

    @Test
    void readAndInsertData() throws Exception {

        // when
        List<Item> coupangItems = itemRepository.findItemsByShop(ShopCode.COUPANG);
        List<Item> naverItems = itemRepository.findItemsByShop(ShopCode.NAVER);

        // then
        Assertions.assertThat(coupangItems.size()).isEqualTo(150);
        Assertions.assertThat(naverItems.size()).isEqualTo(11);

    }
}
