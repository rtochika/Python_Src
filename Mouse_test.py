while True:

    for event in pygame.event.get():
        if event.type == QUIT:
            sys.exit()
        # マウスクリックでカッパをコピー
        if event.type == MOUSEBUTTONDOWN and event.button == 1:
            print("クリックした!!")
            x, y = event.pos
            x -= pythonImg2.get_width() / 2
            y -= pythonImg2.get_height() / 2
            pythons_pos.append((x,y))  # カッパの位置を追加
        # マウス移動でカッパを移動
        if event.type == MOUSEMOTION:
            x, y = event.pos
            x -= pythonImg2.get_width() / 2
            y -= pythonImg2.get_height() / 2
            cur_pos = (x,y)

    # カッパを表示
    screen.blit(pythonImg2, cur_pos)
    for i, j in pythons_pos:
        screen.blit(pythonImg2, (i,j))
    pygame.display.update()