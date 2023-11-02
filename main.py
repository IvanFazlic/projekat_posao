from Functions.Functions import *
connectedSheet = connect_to_sheet()
loadObjects = create_objects_from_excel()
create_the_main_screen(loadObjects, connectedSheet)