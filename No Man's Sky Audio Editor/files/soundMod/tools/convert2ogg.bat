for %%f in (*.stream) do ww2ogg.exe %%f --pcb packed_codebooks_aoTuV_603.bin
for %%f in (*.ogg) do revorb.exe %%f
pause