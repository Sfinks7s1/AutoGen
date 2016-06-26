;;; AutoLisp routine by E.Elovkov
;;; Remove ACAD_DGNLINESTYLECOMP "XRECORD"
(defun-q dictremdgn ()
  (dictremove (namedobjdict) "ACAD_DGNLINESTYLECOMP")
)
(if (dictremdgn)
  (princ
    (strcat
    "\n------------------------"
    "\n Удален объект DICTIONARY \"ACAD_DGNLINESTYLECOMP\""
    "\n Выполните команды \"PURGE\" и \"AUDIT\" "
    )
  ) ;; end princ
  (princ "\n-----------------------\n  Ничего не удалено - объекта \"ACAD_DGNLINESTYLECOMP\" в файле нет")
) ;; end if
(princ)
