# ProfileViewTableNVK
Скрипт для заполнения данных в подпрфильыне таблицы видов профилей Civil 3d.

Для напорных труопроводов заполняются: уклон труб, тип труб из описания трубы из каталога, углы поворота трассы в плане.
Для без напортных труопроводов заполняются: уклон труб, тип труб из описания трубы из каталога.

Константы pipeTypeText, pipeSlopeText, aligAngleText задают префикс имени создаваемых блоков с данным в подпрофильной таблице. К данным прифексам при создании добавляется имя трассы на основании которой был построен вид профиля.\n
k1Yoffset - смещение точки вставки блока уклонов труб от точки вставки вида профиля для видов профилей типа К1.\n
k2Yoffset - смещение точки вставки блока уклонов труб от точки вставки вида профиля для видов профилей типа К2.\n
bYoffset - смещение точки вставки блока уклонов труб от точки вставки вида профиля для видов профилей типа В1 и В2.\n

kTypeOffset - смещение точки вставки блока типов труб от точки вставки вида профиля для видов профилей типа К1 и К2.\n
bTypeOffset - смещение точки вставки блока типов труб от точки вставки вида профиля для видов профилей типа В1 и В2.\n
