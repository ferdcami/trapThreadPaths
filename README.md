eDiscovery Email Threading ID Python Script</br>
***if you have identified thread paths that you want to produce JUST the nodes within the selected path, this python script will take a list, parse the paths, then look them up against a list of paths within your entire list***</br>

Create ParentList.xlsx, Sheet1 called: Sheet1 (contains threading ids to capture-to locate), sheet 2 called: All Docs, contains all records with Email Threading ID populated</br>
Run python_parseThreads.py > generates: thread_permutations.xlsx</br>
Run python_trapThreads.py > generates: ParentList_Flagged.xlsx</br>
