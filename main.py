from view import AppView
from controller import Controller
from config import setup_logger

setup_logger()


if __name__ == "__main__":
    
    app_view = AppView()
    app_view.block_ui()

    # Pour forcer la fenêtre devant les autres
    # app_view.attributes('-topmost', True)
    
    app_view.update()

    controller = Controller(model=None, view=app_view)

    app_view.mainloop()
