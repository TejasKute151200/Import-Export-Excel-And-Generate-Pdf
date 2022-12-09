package com.pdf.repo;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.pdf.entity.PdfEntity;

@Repository
public interface PdfEntityRepo extends JpaRepository<PdfEntity, Integer> {

}
