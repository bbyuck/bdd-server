package com.bb.bdd.domain.excel.entity;

import com.bb.bdd.domain.excel.ShopCode;
import jakarta.persistence.*;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.NoArgsConstructor;

@Table(name = "items")
@Entity
@Getter
@SequenceGenerator(
        name = "item_seq_generator",
        sequenceName = "seq_item",
        allocationSize = 1
)
@NoArgsConstructor(access = AccessLevel.PROTECTED)
public class Item {

    @Id
    @GeneratedValue(strategy = GenerationType.SEQUENCE, generator = "item_seq_generator")
    @Column(name = "id")
    private Long id;

    @Column(name = "source_name")
    private String sourceName;

    @Column(name = "mapping_name")
    private String mappingName;

    @Column(name = "shop")
    private ShopCode shop;
}
