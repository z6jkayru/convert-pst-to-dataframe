from aspose.email.storage.pst import PersonalStorage

# Load PST file
personalStorage = PersonalStorage.from_file("CONVERTED_1.pst")

# Get folders' collection
folderInfoCollection = personalStorage.root_folder.get_sub_folders()

# Extract folders' information
for folderInfo in folderInfoCollection:
    print('')
    print("::::::::::::::::::::::::::::::::::::::::::")
    print("Folder: " + folderInfo.display_name)
    print("Total Items: " + str(folderInfo.content_count))
    print("Total Unread Items: " + str(folderInfo.content_unread_count))

    for subfolderInfo in folderInfo.enumerate_messages():
        ff = personalStorage.get_parent_folder(subfolderInfo.entry_id)
        print(ff.display_name)
        # print("::::::::::::::::::::::::::::::::::::::::::")
        # print("Folder: " + subfolderInfo.display_name)
        # print("Total Items: " + str(subfolderInfo.content_count))
        # print("Total Unread Items: " + str(subfolderInfo.content_unread_count))
