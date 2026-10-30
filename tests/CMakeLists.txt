# The tests of the corpus games, run by WFC: each needs the `wfc` program of the
# WFC build this directory is part of (tests/vb6games/tests of WFC; the
# vb6games CI builds WFC and registers all of them).
#
# The commit gate of WFC registers a smoke pair (Connect 4 played through a
# script, interpreted and compiled). WFC_VB6GAMES_FULL=ON registers the rest:
# the screenshots, the scripted sessions, the Capture King games over sockets
# and the packaged programs.
#
# Variables of WFC used here: WFC_PACKAGE_TESTS (the package test scripts) and
# the target `wfc`.

set(vb6games_dir "${CMAKE_CURRENT_SOURCE_DIR}/..")
set(project_checks "${CMAKE_CURRENT_SOURCE_DIR}")
set(package_tests "${WFC_PACKAGE_TESTS}")

# Registers the screenshot test of one corpus game (vb6games_dir, project_checks
# and package_tests are set above).
# Optional fifth and sixth arguments: a ui-script that drives the windows and
# the screenshot file it must take; an optional seventh, the WFC_RANDOM_SEED of
# the run (a game that plays at random must not end before the screenshot).
function(add_vb6_game name vbp forms milliseconds)
  if(EXISTS "${vb6games_dir}/${vbp}")
    set(script_arguments "")
    if(ARGC GREATER 5)
      set(
        script_arguments
        "-DSCRIPT=${project_checks}/${ARGV4}"
        "-DFILES=${ARGV5}"
      )
    endif()
    if(ARGC GREATER 6)
      list(APPEND script_arguments "-DSEED=${ARGV6}")
    endif()
    # Games run as machine code (MP-0009). SuMBOku 1.4.2 assigns an object
    # without Set (`Dest = Me.Value`), which only the interpreter accepts.
    set(native_argument "-DNATIVE=1")
    if("${name}" STREQUAL "sumboku")
      set(native_argument "-DNATIVE=0")
    endif()
    add_test(
      NAME TC-MP0008-vb6games-${name}
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/${vbp}"
        "-DSHOTS=${CMAKE_BINARY_DIR}/screenshots/${name}" "-DFORMS=${forms}"
        "-DMILLISECONDS=${milliseconds}" ${native_argument} ${script_arguments}
        -P "${project_checks}/run_project_screenshots.cmake"
    )
    set_tests_properties(
      TC-MP0008-vb6games-${name}
      PROPERTIES
        RESOURCE_LOCK "desktop"
        LABELS "integration;windows;mp-0008;vb6games;evidence"
        TIMEOUT 600
    )
  endif()
endfunction()

# Registers a scripted play session of a corpus game: the headless program
# runs the script, whose `get`, `messages` lines state what it expects.
# An optional fifth argument is the WFC_RANDOM_SEED of the session, an optional
# sixth the screen size "WIDTHxHEIGHT" the program sees (WFC_SCREEN), so a
# layout that follows the screen does not depend on the machine's display.
function(add_vb6_play name vbp script expected)
  # The optional seed may be empty, which an unquoted ${ARGN} would drop.
  if(ARGC GREATER 5)
    add_vb6_play_with(
      ""
      ${name}
      ${vbp}
      ${script}
      ${expected}
      "${ARGV4}"
      "${ARGV5}"
    )
  elseif(ARGC GREATER 4)
    add_vb6_play_with("" ${name} ${vbp} ${script} ${expected} "${ARGV4}")
  else()
    add_vb6_play_with("" ${name} ${vbp} ${script} ${expected})
  endif()
endfunction()

# The same session with the program compiled to machine code (MP-0009).
function(add_vb6_native_play name vbp script expected)
  # The optional seed may be empty, which an unquoted ${ARGN} would drop.
  if(ARGC GREATER 5)
    add_vb6_play_with(
      "native"
      ${name}
      ${vbp}
      ${script}
      ${expected}
      "${ARGV4}"
      "${ARGV5}"
    )
  elseif(ARGC GREATER 4)
    add_vb6_play_with("native" ${name} ${vbp} ${script} ${expected} "${ARGV4}")
  else()
    add_vb6_play_with("native" ${name} ${vbp} ${script} ${expected})
  endif()
endfunction()

function(
  add_vb6_play_with
  mode
  name
  vbp
  script
  expected
)
  set(native_argument "")
  set(test_name TC-MP0008-vb6games-${name})
  set(test_labels "integration;mp-0008;vb6games")
  if(NOT "${mode}" STREQUAL "")
    set(native_argument "-DNATIVE=1")
    set(test_name TC-MP0009-native-vb6games-${name})
    set(test_labels "integration;mp-0009;native;vb6games")
  endif()
  if(EXISTS "${vb6games_dir}/${vbp}")
    set(seed_argument "")
    if(ARGC GREATER 5 AND NOT "${ARGV5}" STREQUAL "")
      set(seed_argument "-DSEED=${ARGV5}")
    endif()
    set(screen_argument "")
    if(ARGC GREATER 6)
      set(screen_argument "-DSCREEN=${ARGV6}")
    endif()
    add_test(
      NAME ${test_name}
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/${vbp}"
        "-DSCRIPT=${project_checks}/${script}"
        "-DEXPECTED=${project_checks}/${expected}"
        "-DWORK=${CMAKE_BINARY_DIR}/vb6games-${mode}${name}" ${seed_argument}
        ${screen_argument} ${native_argument} -P
        "${project_checks}/run_project_script.cmake"
    )
    set_tests_properties(
      ${test_name}
      PROPERTIES LABELS "${test_labels}" TIMEOUT 300
    )
    # Capture King listens on a UDP port.
    if("${name}" MATCHES "captureking")
      set_tests_properties(${test_name} PROPERTIES RESOURCE_LOCK udp)
    endif()
  endif()
endfunction()

# Registers the packaged-program test of one corpus game: `wfc --build` makes
# a standalone program, which runs on its own and shows the named form. An
# optional fourth argument is the WFC_RANDOM_SEED of the run.
function(add_vb6_package name vbp form)
  if(EXISTS "${vb6games_dir}/${vbp}")
    set(seed_argument "")
    if(ARGC GREATER 3)
      set(seed_argument "-DSEED=${ARGV3}")
    endif()
    add_test(
      NAME TC-MP0008-package-${name}
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/${vbp}"
        "-DWORK=${CMAKE_BINARY_DIR}/package-${name}" "-DFORM=${form}"
        ${seed_argument} -P
        "${package_tests}/run_package_gui.cmake"
    )
    set_tests_properties(
      TC-MP0008-package-${name}
      PROPERTIES
        RESOURCE_LOCK "desktop"
        LABELS "integration;windows;mp-0008;package"
        TIMEOUT 600
    )
  endif()
endfunction()

# The same packaged-program test with the program compiled to machine code
# (MP-0009): the packaged program carries the code, not a program to read.
function(add_vb6_native_package name vbp form)
  if(EXISTS "${vb6games_dir}/${vbp}")
    set(seed_argument "")
    if(ARGC GREATER 3)
      set(seed_argument "-DSEED=${ARGV3}")
    endif()
    add_test(
      NAME TC-MP0009-native-package-${name}
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/${vbp}"
        "-DWORK=${CMAKE_BINARY_DIR}/package-native-${name}" "-DFORM=${form}"
        "-DNATIVE=1" ${seed_argument} -P
        "${package_tests}/run_package_gui.cmake"
    )
    set_tests_properties(
      TC-MP0009-native-package-${name}
      PROPERTIES
        RESOURCE_LOCK "desktop"
        LABELS "integration;windows;mp-0009;native;package"
        TIMEOUT 600
    )
  endif()
endfunction()

# The same with the program linked dynamically (MP-0010): a small launcher
# program that loads wfcrt.dll, bundled beside it.
function(add_vb6_dynamic_package name vbp form)
  if(EXISTS "${vb6games_dir}/${vbp}")
    set(seed_argument "")
    if(ARGC GREATER 3)
      set(seed_argument "-DSEED=${ARGV3}")
    endif()
    add_test(
      NAME TC-MP0010-dynamic-package-${name}
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/${vbp}"
        "-DWORK=${CMAKE_BINARY_DIR}/package-dynamic-${name}" "-DFORM=${form}"
        "-DNATIVE=1" "-DLINK=dynamic" ${seed_argument} -P
        "${package_tests}/run_package_gui.cmake"
    )
    set_tests_properties(
      TC-MP0010-dynamic-package-${name}
      PROPERTIES
        RESOURCE_LOCK "desktop"
        LABELS "integration;windows;mp-0010;native;package"
        TIMEOUT 600
    )
  endif()
endfunction()

# --- Smoke tests (always) ------------------------------------------------
add_vb6_play(
  connect4-play
  connect4/PConnect4.vbp
  connect4-play.ui
  connect4-play.expected
  ""
  2560x1440
)
add_vb6_native_play(
  connect4-play
  connect4/PConnect4.vbp
  connect4-play.ui
  connect4-play.expected
  ""
  2560x1440
)

# --- The rest (WFC_VB6GAMES_FULL) ----------------------------------------
if(WFC_VB6GAMES_FULL)
  if(EXISTS "${vb6games_dir}/captureking")
    add_test(
      NAME TC-MP0008-vb6games-captureking
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
        "-DSHOTS=${CMAKE_BINARY_DIR}/screenshots/captureking"
        "-DFORMS=frmSplash;frmGame" "-DMILLISECONDS=12000"
        "-DSCRIPT=${project_checks}/captureking-gui.ui"
        "-DFILES=frmGame-moves.png" -P
        "${project_checks}/run_project_screenshots.cmake"
    )
    add_test(
      NAME TC-MP0008-vb6games-captureking-script
      COMMAND
        "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
        "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
        "-DSCRIPT=${project_checks}/captureking.ui"
        "-DEXPECTED=${project_checks}/captureking.expected"
        "-DWORK=${CMAKE_BINARY_DIR}/vb6games-captureking-script" -P
        "${project_checks}/run_project_script.cmake"
    )
    set_tests_properties(
      TC-MP0008-vb6games-captureking-script
      PROPERTIES
        RESOURCE_LOCK udp
        LABELS "integration;mp-0008;vb6games"
        TIMEOUT 60
    )
    # Complete games of random moves and chat messages, played until a king
    # is taken and check mate is called; the saved board must match.
    foreach(game_seed 20261012 99)
      add_test(
        NAME TC-MP0008-vb6games-captureking-game-${game_seed}
        COMMAND
          "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
          "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
          "-DSEED=${game_seed}"
          "-DWORK=${CMAKE_BINARY_DIR}/vb6games-captureking-game-${game_seed}" -P
          "${project_checks}/run_captureking_game.cmake"
      )
      set_tests_properties(
        TC-MP0008-vb6games-captureking-game-${game_seed}
        PROPERTIES
          RESOURCE_LOCK udp
          LABELS "integration;mp-0008;vb6games"
          TIMEOUT 120
      )
    endforeach()
    # Two programs on a virtual LAN (own loopback address and user name each)
    # play a game over real sockets: invite, accept, moves, chat, check mate.
    foreach(lan_seed 99 20261012)
      add_test(
        NAME TC-MP0008-vb6games-captureking-lan-${lan_seed}
        COMMAND
          "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
          "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
          "-DSEED=${lan_seed}"
          "-DWORK=${CMAKE_BINARY_DIR}/vb6games-captureking-lan-${lan_seed}" -P
          "${project_checks}/run_captureking_lan_game.cmake"
      )
      set_tests_properties(
        TC-MP0008-vb6games-captureking-lan-${lan_seed}
        PROPERTIES
          RESOURCE_LOCK udp
          LABELS "integration;mp-0008;vb6games;network"
          TIMEOUT 300
      )
    endforeach()
    set_tests_properties(
      TC-MP0008-vb6games-captureking
      PROPERTIES
        RESOURCE_LOCK "desktop;udp"
        LABELS "integration;windows;mp-0008;vb6games;evidence"
        TIMEOUT 90
    )
  endif()
  # Further corpus games: each runs in real windows; the forms it shows must
  # each leave a screenshot.
  add_vb6_game(clix clix/src/Clix.vbp "frmMain" 5000)
  add_vb6_play(
    restaum-solve
    restaum/source/BrainVita.vbp
    restaum-solve.ui
    restaum-solve.expected
  )
  add_vb6_play(
    sumboku-solve
    sumboku/PSuMBOku.vbp
    sumboku-solve.ui
    sumboku-solve.expected
  )
  add_vb6_play(
    executioner-play
    executioner/Executioner.vbp
    executioner-play.ui
    executioner-play.expected
    1
  )
  add_vb6_game(connect4 connect4/PConnect4.vbp "FMain" 5000)
  add_vb6_game(sumboku-current sumboku/PSuMBOku.vbp "FrmSudoku" 5000)
  add_vb6_game(
    executioner
    executioner/Executioner.vbp
    "frmSplash;frmGamescreen"
    6000
  )
  add_vb6_game(
    pipes
    pipes/pipes.vbp
    "Logo"
    30000
    pipes-gui.ui
    Form1-board.png
    2
  )
  add_vb6_play(pipes-play pipes/pipes.vbp pipes-play.ui pipes-play.expected 1)
  # The same sessions with the programs compiled to machine code (MP-0009).
  add_vb6_native_play(
    restaum-solve
    restaum/source/BrainVita.vbp
    restaum-solve.ui
    restaum-solve.expected
  )
  add_vb6_native_play(
    executioner-play
    executioner/Executioner.vbp
    executioner-play.ui
    executioner-play.expected
    1
  )
  add_vb6_native_play(
    pipes-play
    pipes/pipes.vbp
    pipes-play.ui
    pipes-play.expected
    1
  )
  add_vb6_native_play(
    sumboku-solve
    sumboku/PSuMBOku.vbp
    sumboku-solve.ui
    sumboku-solve.expected
  )
  add_vb6_native_play(
    captureking-script
    captureking/CaptureKing.vbp
    captureking.ui
    captureking.expected
  )
  add_vb6_game(restaum restaum/source/BrainVita.vbp "Form1" 5000)
  add_vb6_package(clix clix/src/Clix.vbp frmMain 3)
  add_vb6_package(connect4 connect4/PConnect4.vbp FMain)
  add_vb6_package(executioner executioner/Executioner.vbp frmGamescreen)
  add_vb6_package(pipes pipes/pipes.vbp Logo)
  add_vb6_package(restaum restaum/source/BrainVita.vbp Form1)
  add_vb6_native_package(clix clix/src/Clix.vbp frmMain 3)
  add_vb6_native_package(connect4 connect4/PConnect4.vbp FMain)
  add_vb6_dynamic_package(connect4 connect4/PConnect4.vbp FMain)
  add_vb6_dynamic_package(pipes pipes/pipes.vbp Logo)
  add_vb6_native_package(executioner executioner/Executioner.vbp frmGamescreen)
  add_vb6_native_package(pipes pipes/pipes.vbp Logo)
  add_vb6_native_package(restaum restaum/source/BrainVita.vbp Form1)
  add_vb6_native_package(sumboku-current sumboku/PSuMBOku.vbp FrmSudoku)
  add_vb6_package(sumboku sumboku/v1.4.2/SuMBOku_V_1_4/Sudoku.vbp FrmSudoku)
  add_vb6_game(sumboku sumboku/v1.4.2/SuMBOku_V_1_4/Sudoku.vbp "FrmSudoku" 5000)
  if(WIN32)
    if(EXISTS "${vb6games_dir}/captureking")
      add_test(
        NAME TC-MP0008-package-icon
        COMMAND
          "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
          "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
          "-DWORK=${CMAKE_CURRENT_BINARY_DIR}/package-icon"
          "-DCHECK=${package_tests}/icon_differs.ps1" -P
          "${package_tests}/run_icon_test.cmake"
      )
      set_tests_properties(
        TC-MP0008-package-icon
        PROPERTIES LABELS "integration;mp-0008;package"
      )
      add_test(
        NAME TC-MP0008-package-standalone-gui
        COMMAND
          "${CMAKE_COMMAND}" "-DWFC=$<TARGET_FILE:wfc>"
          "-DPROJECT=${vb6games_dir}/captureking/CaptureKing.vbp"
          "-DWORK=${CMAKE_CURRENT_BINARY_DIR}/package-standalone"
          "-DFORM=frmGame" -P "${package_tests}/run_package_gui.cmake"
      )
      set_tests_properties(
        TC-MP0008-package-standalone-gui
        PROPERTIES
          RESOURCE_LOCK "desktop;udp"
          LABELS "integration;windows;mp-0008;package"
          TIMEOUT 60
      )
    endif()
  endif()
endif()
